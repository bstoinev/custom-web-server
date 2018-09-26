using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using PortSys.Tac.ClientServices.Kernel.ExtensionMethods.IEnumerable.Win32;

namespace PortSys.Tac.ClientServices.Kernel.Win32
{
    public class LaunchProcessCommandWin32Helper : Win32Helper
    {
        #region Win32

        #region Const

        protected const uint ZERO_FLAGS = 0;
        
        protected IntPtr WTS_CURRENT_SERVER_HANDLE = IntPtr.Zero;

        protected enum WTS_CONNECTSTATE_CLASS
        {
            WTSActive,
            WTSConnected,
            WTSConnectQuery,
            WTSShadow,
            WTSDisconnected,
            WTSIdle,
            WTSListen,
            WTSReset,
            WTSDown,
            WTSInit
        }

        #endregion

        #region Struct

        public struct SECURITY_ATTRIBUTES
        {
            public int Size;
            public IntPtr SecurityDescriptorPointer;
            public bool EnableHandleInheritance;
        }


        public struct PROCESS_INFORMATION
        {
            public IntPtr ProcessHandle;
            public IntPtr ThreadHandle;
            public uint ProcessId;
            public uint ThreadId;
        }

        protected struct WTS_SESSION_INFO
        {
            public Int32 SessionID;
            [MarshalAs(UnmanagedType.LPStr)]
            public String WinStationName;
            public WTS_CONNECTSTATE_CLASS State;
        }

        #endregion

        [DllImport("wtsapi32.dll", SetLastError = false)]
        public extern static void WTSFreeMemory(IntPtr ObjectPointer);

        /// <summary>
        /// [MSDN] Retrieves a list of sessions on a specified Remote Desktop Session Host (RD Session Host) server.
        /// </summary>
        /// <param name="ServerHandle"></param>
        /// <param name="Reserved">Always 0</param>
        /// <param name="Version">Always 1</param>
        /// <param name="ppSessionInfo">Pointer to array of WTS_SESSION_INFO</param>
        /// <param name="pCount">Pointer to number of WTS_SESSION_INFO structures.</param>
        /// <see cref="http://msdn.microsoft.com/en-us/library/aa383833(v=vs.85).aspx"/>
        [DllImport("wtsapi32.dll", SetLastError = true)]
        static extern bool WTSEnumerateSessions(
            [In] IntPtr ServerHandle,
            [In] uint Reserved,
            [In] uint Version,
            [Out] out IntPtr SessionInfoPointer,
            [Out] out int SessionInfoLength
        );
        
        /// <summary>
        /// [MSDN] Retrieves the Remote Desktop Services session that is currently attached to the physical console.
        /// </summary>
        /// <see cref="http://msdn.microsoft.com/en-us/library/aa383835(v=vs.85).aspx"/>
        [DllImport("kernel32.dll")]
        static extern uint WTSGetActiveConsoleSessionId();

        [DllImport("wtsapi32.dll", SetLastError = true)]
        static extern bool WTSQueryUserToken(uint SessionId, out IntPtr Token);

        /// <summary>
        /// [MSDN] The ImpersonateLoggedOnUser function lets the calling thread impersonate the security context of a logged-on user.
        /// </summary>
        /// <param name="UserAccessToken">Handle to a logged-on user's access token.</param>
        /// <see cref="http://msdn.microsoft.com/en-us/library/windows/desktop/aa378612(v=vs.85).aspx" />
        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool ImpersonateLoggedOnUser([In] IntPtr UserAccessToken);

        /// <summary>
        /// [MSDN] Creates a new process and its primary thread. 
        /// The new process runs in the security context of the user represented by the specified token.
        /// </summary>
        /// <param name="Token">Handle to the access token, under which the process and its primary thread starts.</param>
        /// <param name="ApplicationPathFile">Executable name, including extension. Null if the filename is specified in <paramref name="CommandLineArguments"/></param>
        /// <param name="CommandLineArguments">The command line.</param>
        /// <param name="ProcessSecurityAttributes">Indicates the security descriptor for the new process.</param>
        /// <param name="ThreadSecurityAttributes">Indicates the security descriptor for the new process' primary thread.</param>
        /// <param name="EnableHandleInheritance">true, to inherit handles from the calling process; otherwise false.</param>
        /// <param name="CreationFlags">Indicates the process priority, window state, etc.</param>
        /// <param name="Environment">[MSDN] A pointer to an environment block for the new process. If this parameter is NULL, the new process uses the environment of the calling process. </param>
        /// <param name="CurrentDirectory">The full path to the working directory of the process.</param>
        /// <param name="StartupInfo">Pointer to a structure, which specifies the desktop, window station, etc.</param>
        /// <param name="ProcessInformation">Pointer to a structure, that receives identification information of the newly created process.</param>
        /// <see cref="http://msdn.microsoft.com/en-us/library/windows/desktop/ms682429(v=vs.85).aspx" />
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern bool CreateProcessAsUser(
            IntPtr Token,
            string ApplicationPathFile,
            string CommandLineArguments,
            ref SECURITY_ATTRIBUTES ProcessSecurityAttributes,  
            ref SECURITY_ATTRIBUTES ThreadSecurityAttributes,   
            bool EnableHandleInheritance,
            uint CreationFlags,
            IntPtr Environment,
            string CurrentDirectory,
            ref STARTUPINFO StartupInfo, 
            out PROCESS_INFORMATION ProcessInformation
        );

        /// <summary>
        /// [MSDN] Performs an operation on a specified file.
        /// </summary>
        /// <param name="ParentWindow">Handle to a parent window. null for startup operations.</param>
        /// <param name="Verb">`Specified the action to be performed. E.g. "edit", "find", "open", etc.</param>
        /// <param name="PathFile">FQFN to the executable file of the program OR the file, associated with a program.</param>
        /// <param name="CommandArguments">
        /// If <paramref name="PathFile" /> specifies an executable file, <paramref name="CommandArguments"/> specifies command line parameters, if any.</param>
        /// If <paramref name="PathFile" /> specifies a document, this parameter should be null.
        /// <param name="WorkingDirectory">null, to use the current working directory.</param>
        /// <param name="ShowCommand">Flags, that determines how tha pplication is displayed when it is opened.</param>
        /// <returns>
        /// After converting it to Int32, the return value is number greater than 32, if the function succeeds.
        /// If it is less than 33, an error occured.
        /// </returns>
        /// <see cref="http://msdn.microsoft.com/en-us/library/windows/desktop/bb762153(v=vs.85).aspx" />
        [DllImport("shell32.dll")]
        static extern IntPtr ShellExecute(
            IntPtr ParentWindow,
            string Verb,
            string PathFile,
            string CommandArguments,
            string WorkingDirectory,
            int ShowCommand
        );

        #endregion
        
        public LaunchProcessCommandWin32Helper() : base(1)
        {
            StartupInfo = STARTUPINFO.DefaultStartupInfo();
        }

        public STARTUPINFO? StartupInfo { get; set; }
        
        private string LookupOpenVerb(string FileExtension)
        {
            string result = null;

            var fileTypeAssoc = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(FileExtension);
            if (fileTypeAssoc != null)
            {
                var progId = fileTypeAssoc.GetValue(null) ?? string.Empty;
                var rkOpenCommand = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(string.Format("{0}\\shell\\open\\command", progId));
                if (rkOpenCommand != null)
                {
                    result = rkOpenCommand.GetValue(null) as string;
                }
            }

            return result;
        }

        public void SpawnProcessToActiveConsole(string TargetFile, string CommandArguments, string WorkingFolder)
        {
            var siPointer = IntPtr.Zero;
            var siCount = 0;
            var siSize = Marshal.SizeOf(typeof(WTS_SESSION_INFO));

            if (!WTSEnumerateSessions(WTS_CURRENT_SERVER_HANDLE, 0, 1, out siPointer, out siCount))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Failed to enumerate sessions.");
            }

            var siList = new List<WTS_SESSION_INFO>((int)siCount);
            siList.ReadStructureArray(siPointer, siCount);

            WTSFreeMemory(siPointer);

            var sessionId = uint.MaxValue;
            foreach (var si in siList)
            {
                if (si.State == WTS_CONNECTSTATE_CLASS.WTSActive)
                {
                    if (sessionId != uint.MaxValue)
                    {
                        throw new NotSupportedException("Multiple desktops detected.\nThe LaunchProcessCommand is not supported on multidesktop system.");
                    }
                    sessionId = (uint)si.SessionID;
                }
            }

            if (sessionId == uint.MaxValue)
            {
                throw new InvalidOperationException("No active session found.");
            }

            var commandFile = TargetFile.Trim('"');
            var fileExt = Path.GetExtension(commandFile);
            commandFile = fileExt.Equals(".exe", StringComparison.OrdinalIgnoreCase) ? commandFile : LookupOpenVerb(fileExt);
            
            if (string.IsNullOrEmpty(commandFile))
            {
                throw new InvalidOperationException("Unable to locate file type handler.");
            }

            var cmdArgs = CommandArguments;
            var workingDirectory = WorkingFolder ?? Path.GetDirectoryName(commandFile);
            var hAccessToken = CreateHandle();
            if (!WTSQueryUserToken(sessionId, out hAccessToken))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Failed to query user's access token.");
            }

            var secAttribs = new SECURITY_ATTRIBUTES();
            var creationFlags = ZERO_FLAGS;
            var envBlock = IntPtr.Zero;
            var startInfo = StartupInfo ?? STARTUPINFO.DefaultStartupInfo();
            var procInfo = new PROCESS_INFORMATION();
            var launched = CreateProcessAsUser(
                hAccessToken,
                commandFile,
                cmdArgs,
                ref secAttribs,
                ref secAttribs,
                false,
                creationFlags,
                envBlock,
                workingDirectory,
                ref startInfo,
                out procInfo
            );
                        
            if (!launched)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), string.Format("Unable to execute the following command:\n{0} {1}", commandFile, cmdArgs));
            }
        }
    }
}
