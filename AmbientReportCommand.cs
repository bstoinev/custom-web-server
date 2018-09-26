using NetFwTypeLib;
using PortSys.Tac.ClientServices.Intelligence;
using PortSys.Tac.ClientServices.Kernel.ExtensionMethods.IEnumerable;
using PortSys.Tac.ClientServices.Kernel.ExtensionMethods.ManagementObjectCollection;
using System;
using System.Management;

namespace PortSys.Tac.ClientServices.Kernel.Processing
{
    [CommandName("AmbientReport")]
    public sealed class AmbientReportCommand : ICommand<ICommandParameters, AmbientReportResult>
    {
        private Version VistaSp1Version = new Version(6, 0, 6001, 18000);

        public AmbientReportCommand()
        {
            State = CommandState.NotStarted;
            Monitor = new CommandMonitor("EngineEventSource");
        }

        public CommandState State { get; private set; }

        public ICommandParameters Parameters
        {
            get { return null; }
        }

        public AmbientReportResult Result { get; private set; }

        public IMonitor Monitor { get; private set; }

        private void LoadWscAntivirusProductInfo()
        {
            var useLatestWsc = Environment.OSVersion.Version >= VistaSp1Version;
            var scope = string.Format(@"root\{0}", useLatestWsc ? "SecurityCenter2" : "SecurityCenter");
            var searcher = new ManagementObjectSearcher(scope, "SELECT * FROM AntivirusProduct");
            var wmiResponse = searcher.Get().ToList();

            if (wmiResponse.Count == 0)
            {
                Monitor.Debug("No antivirus product is installed.");
                Result.Message += "Antivirus product not installed.\n";
            }
            else
            {
                if (wmiResponse.Count != 1)
                {
                    var strangeResult = string.Format("Security Center reported more than one antivirus products are installed [Count = {0}]. Only the first one is included in the result.", wmiResponse.Count);

                    Monitor.Warning("Unexpected WMI response is encountered.");
                    Monitor.Debug(strangeResult);
                }

                var foundAvProduct = wmiResponse.FirstOrDefault<ManagementBaseObject>();

                var enabledPropertyName = useLatestWsc ? "productState" : "onAccessScanningEnabled";
                var updatedPropertyName = useLatestWsc ? "productState" : "productUptoDate";

                var displayNameProperty = foundAvProduct.Properties.FirstOrDefault<PropertyData>(pd => pd.Name.Equals("displayName", StringComparison.OrdinalIgnoreCase));
                var enabledProperty = foundAvProduct.Properties.FirstOrDefault<PropertyData>(pd => pd.Name.Equals(enabledPropertyName, StringComparison.OrdinalIgnoreCase));
                var updatedProperty = foundAvProduct.Properties.FirstOrDefault<PropertyData>(pd => pd.Name.Equals(updatedPropertyName, StringComparison.OrdinalIgnoreCase));

                if (displayNameProperty == null || enabledProperty == null || updatedProperty == null)
                {
                    Monitor.Error("Antivirus product info could not be retrieved.");
                    Result.Message += "Antivirus product state could not be retrieved.";
                }
                else
                {
                    var enabledPropertyValue = useLatestWsc ? DecodeWsc2ProductEnabledState(enabledProperty) : Convert.ToBoolean(enabledProperty.Value);
                    var updatedPropertyValue = useLatestWsc ? DecodeWsc2ProductUpdatedState(updatedProperty) : Convert.ToBoolean(updatedProperty.Value);

                    Result.Antivirus = new AmbientReportAntivirusInfo();
                    Result.Antivirus.DisplayName = displayNameProperty.Value.ToString();
                    Result.Antivirus.Enabled = enabledPropertyValue;
                    Result.Antivirus.Updated = updatedPropertyValue;
                }
            }
        }

        private void LoadWscFirewallProductInfo()
        {
            var useLatestWsc = Environment.OSVersion.Version >= VistaSp1Version;
            var scope = string.Format(@"root\{0}", useLatestWsc ? "SecurityCenter2" : "SecurityCenter");
            var searcher = new ManagementObjectSearcher(scope, "SELECT * FROM FirewallProduct");
            var wmiResponse = searcher.Get().ToList();

            if (wmiResponse.Count == 0)
            {
                Monitor.Debug("No third party firewall product is installed.");
                Result.Message += "Third party firewall product not installed.\n";
            }
            else
            {
                if (wmiResponse.Count != 1)
                {
                    var strangeResult = string.Format("Security Center reported more than one firewall products are installed [Count = {0}]. Only the first one is included in the result.", wmiResponse.Count);

                    Monitor.Warning("Unexpected WMI response is encountered.");
                    Monitor.Debug(strangeResult);
                }

                var foundFirewallProduct = wmiResponse.FirstOrDefault<ManagementBaseObject>();

                var enabledPropertyName = useLatestWsc ? "productState" : "onAccessScanningEnabled";

                var displayNameProperty = foundFirewallProduct.Properties.FirstOrDefault<PropertyData>(pd => pd.Name.Equals("displayName", StringComparison.OrdinalIgnoreCase));
                var enabledProperty = foundFirewallProduct.Properties.FirstOrDefault<PropertyData>(pd => pd.Name.Equals(enabledPropertyName, StringComparison.OrdinalIgnoreCase));

                if (displayNameProperty == null || enabledProperty == null)
                {
                    Monitor.Debug("Firewall product info could not be retrieved.");
                    Result.Message += "Third party firewall state could not be retrieved.\n";
                }
                else
                {
                    var enabledPropertyValue = useLatestWsc ? DecodeWsc2ProductEnabledState(enabledProperty) : Convert.ToBoolean(enabledProperty.Value);                    

                    Result.Firewall = new AmbientReportFirewallInfo();
                    Result.Firewall.DisplayName = displayNameProperty.Value.ToString();
                    Result.Firewall.Enabled = enabledPropertyValue;
                }
            }
        }

        private bool DecodeWsc2ProductEnabledState(PropertyData ProductState)
        {
            var maskHex = Convert.ToInt32(ProductState.Value).ToString("X6");
            var enabledState = byte.Parse(maskHex.Substring(2, 1)); // WSC_SECURITY_PRODUCT_STATE enumeration

            return enabledState == 1;
        }

        private bool DecodeWsc2ProductUpdatedState(PropertyData ProductState)
        {
            var maskHex = Convert.ToInt32(ProductState.Value).ToString("X6");
            var updatedState = byte.Parse(maskHex.Substring(4, 2)); // WSC_SECURITY_SIGNATURE_STATUS enumeration

            return updatedState == 0;
        }

        private AmbientReportWindowsFirewallInfo GetWindowsFirewallManagerSettings()
        {
            AmbientReportWindowsFirewallInfo result = null;

            INetFwMgr firewallManager = null;

            try
            {
                firewallManager = (INetFwMgr)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwMgr"));
            }
            catch (Exception ex)
            {
                var errorMessage = string.Format("Failed to instantiate Windows Firewall manager with the following exception:\n{1}", ex.ToString());
                Monitor.Error(errorMessage);
            }

            if (firewallManager != null)
            {
                result = new AmbientReportWindowsFirewallInfo();

                switch (firewallManager.CurrentProfileType)
                {
                    case NET_FW_PROFILE_TYPE_.NET_FW_PROFILE_DOMAIN:
                        result.DomainEnabled = firewallManager.LocalPolicy.CurrentProfile.FirewallEnabled;
                        break;
                    case NET_FW_PROFILE_TYPE_.NET_FW_PROFILE_STANDARD:
                        result.StandardEnabled = firewallManager.LocalPolicy.CurrentProfile.FirewallEnabled;
                        break;
                    case NET_FW_PROFILE_TYPE_.NET_FW_PROFILE_TYPE_MAX:
                    case NET_FW_PROFILE_TYPE_.NET_FW_PROFILE_CURRENT:
                    default:
                        Monitor.Warning("Unknown Windows Firewall profile type: {0}", firewallManager.CurrentProfileType);
                        Result.Message += "Unknown Windows Firewall state.";
                        break;
                }
            }

            return result;
        }

        private AmbientReportWindowsFirewallInfo GetWindowsFirewallPolicySettings()
        {
            AmbientReportWindowsFirewallInfo result = null;

            INetFwPolicy2 firewallPolicy = null;

            try
            {
                firewallPolicy = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
            }
            catch (Exception ex)
            {
                var errorMessage = string.Format("Failed to instantiate Windows firewall policy with the following exception:\n{1}", ex.ToString());
                Monitor.Error(errorMessage);
            }

            if (firewallPolicy != null)
            {
                var currentProfile = (NET_FW_PROFILE_TYPE2_)firewallPolicy.CurrentProfileTypes;

                if (currentProfile == NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_ALL)
                {
                    var warnMessage = string.Format("Inconclusive Windows firewall profile type: {1}", currentProfile);
                    Monitor.Warning(warnMessage);
                }
                else
                {
                    result = new AmbientReportWindowsFirewallInfo();

                    if ((currentProfile & NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_DOMAIN) == NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_DOMAIN)
                    {
                        result.DomainEnabled = firewallPolicy.FirewallEnabled[NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_DOMAIN];
                    }

                    if ((currentProfile & NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_PRIVATE) == NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_PRIVATE)
                    {
                        result.PrivateEnabled = firewallPolicy.FirewallEnabled[NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_PRIVATE];
                    }

                    if ((currentProfile & NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_PUBLIC) == NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_PUBLIC)
                    {
                        result.PublicEnabled = firewallPolicy.FirewallEnabled[NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_PUBLIC];
                    }
                }
            }

            return result;
        }

        private void LoadWindowsFirewallInfo()
        {
            Result.WindowsFirewall = Environment.OSVersion.Version.Major < 6 ? GetWindowsFirewallManagerSettings() : GetWindowsFirewallPolicySettings();

            if (Result.WindowsFirewall == null)
            {
                Monitor.Warning("Unable to determine Windows firewall state.");
                Result.Message += "Windows firewall state cannot be determined.\n";
            }
        }

        public void ReadParameters(string JsonString) { }

        public void Run()
        {
            State = CommandState.Running;

            Result = new AmbientReportResult();

            Monitor.Debug("Retrieving OS info...");
            LoadOsInfo();
            Monitor.Debug("Retrieving Windows firewall info...");
            LoadWindowsFirewallInfo();
            Monitor.Debug("Retrieving antivirus info...");
            LoadWscAntivirusProductInfo();
            Monitor.Debug("Retrieving third party firewall info...");
            LoadWscFirewallProductInfo();

            Result.Message += "Command completed.";
            State = CommandState.Completed;
        }

        private void LoadOsInfo()
        {
            Result.OS = new AmbientReportOsInfo();
            Result.OS.DisplayName = Environment.OSVersion.VersionString;
            Result.OS.Version = Environment.OSVersion.Version;
            Result.OS.Platform = Environment.OSVersion.Platform.ToString();
        }
    }
}
