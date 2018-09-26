using System;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Diagnostics;
using System.Timers;

namespace PortSys.Tac.ClientServices.Hosting
{
    public class KernelBootstrapper : MarshalByRefObject, IDisposable
    {
        private TraceSource Log = new TraceSource("HostEventSource");

        private bool Rebooting = false;

        internal AppDomain KernelPartition { get; private set; }
        internal ITcsKernel KernelTurbine { get; private set; }

        #region Restart second

        private Timer RestartSecond;

        private void InitializeSecond()
        {
            int retryTimeout = 30;
            int.TryParse(ConfigurationManager.AppSettings["startupSequenceRetryIntervalInSeconds"], out retryTimeout);
            RestartSecond = new Timer(retryTimeout * 1000);
            RestartSecond.Elapsed += RestartSecond_TimeElapsed;
        }

        private void DisposeKernel()
        {
            if (KernelPartition != null)
            {
                if (KernelTurbine != null)
                {
                    KernelTurbine.Dispose();
                    KernelTurbine = null;
                }

                KernelPartition.DomainUnload -= KernelPartition_DomainUnload;
                AppDomain.Unload(KernelPartition);
            }
        }

        private void DisposeRestartSecond()
        {
            if (RestartSecond != null)
            {
                RestartSecond.Dispose();
            }
        }

        private void RestartSecond_TimeElapsed(object s, ElapsedEventArgs e)
        {
            RestartSecond.Enabled = false;
            Reboot();
        }

        #endregion

        private void KernelPartition_DomainUnload(object sender, EventArgs e)
        {
            KernelPartition = null;
            if (Rebooting)
            {
                Rebooting = false;
                Boot();
            }
        }

        private bool ValidateConfig()
        {
            var cfg = TacCsConfigSection.LoadConfig(true);
            if (cfg == null)
            {
                Log.TraceEvent(TraceEventType.Warning, 0, "Configuration is either missing or corrupt.");
            }

            return cfg != null;
        }

        private void InitializeRestart()
        {
            if (RestartSecond == null)
            {
                InitializeSecond();
            }

            Log.TraceInformation(string.Format("Attempting restart in {0} seconds.", RestartSecond.Interval / 1000));
            RestartSecond.Enabled = true;
        }

        private void Start()
        {
            if (ValidateConfig())
            {
                KernelTurbine.Run();
            }
            else
            {
                Log.TraceEvent(TraceEventType.Error, 0, "Cannot load start without configuration. Fix the config file and try again.");
            }
        }

        public override object InitializeLifetimeService()
        {
            return null;
        }

        public void Reboot()
        {
            if (KernelTurbine != null)
            {
                KernelTurbine.Dispose();
                KernelTurbine = null;
            }

            if (KernelPartition == null)
            {
                Boot();
            }
            else
            {
                Rebooting = true;
                AppDomain.Unload(KernelPartition);
            }
        }

        public void Boot()
        {
            if (KernelPartition != null)
            {
                throw new InvalidOperationException("Kernel partition already exists.");
            }

            Log.TraceEvent(TraceEventType.Verbose, 0, "Initiating startup sequence...");
            KernelPartition = AppDomain.CreateDomain("TAC Client Services Kernel", null, new AppDomainSetup() { ApplicationName = "krnl", ShadowCopyFiles = "true" });
            KernelPartition.DomainUnload += KernelPartition_DomainUnload;

            string engineClassName = null;
            string engineAssemblyName = null;

            try
            {
                var engineTypeId = ConfigurationManager.AppSettings["engineTypeId"];
                engineClassName = engineTypeId.Split(',')[0].Trim();
                engineAssemblyName = engineTypeId.Substring(engineClassName.Length + 1).Trim();
            }
            catch (Exception)
            {
                Log.TraceEvent(TraceEventType.Verbose, 0, "Configuration errors detected. Attempting defaults...");
                engineClassName = "PortSys.Tac.ClientServices.Kernel.Turbine";
                engineAssemblyName = "PortSys.Tac.ClientServices.Kernel";
            }
            
            try
            {
                KernelTurbine = (ITcsKernel)KernelPartition.CreateInstanceAndUnwrap(engineAssemblyName, engineClassName);
                Log.TraceEvent(TraceEventType.Verbose, 0, "Kernel connection established.");
            }
            catch (Exception ex)
            {
                Log.TraceEvent(TraceEventType.Verbose, 0, "Failed to establish communication channel with the kernel with the following exception:\n{0}", ex.ToString());
            }

            if (KernelTurbine == null)
            {
                InitializeRestart();
            }
            else
            {
                try
                {
                    Start();
                    Log.TraceEvent(TraceEventType.Verbose, 0, "Startup sequence completed.");
                    KernelTurbine.RebootDelegate = Reboot;
                    DisposeRestartSecond();
                }
                catch (Exception ex)
                {
                    string message = string.Format("The startup sequence failed with the following exception:{0}{1}", Environment.NewLine, ex.ToString());
                    Log.TraceEvent(TraceEventType.Error, 0, message);

                    InitializeRestart();
                }
            }
        }

        public void Dispose()
        {
            DisposeRestartSecond();
            DisposeKernel();
        }
    }
}
