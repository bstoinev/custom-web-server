using PortSys.Tac.ClientServices.Intelligence;
using System;
using System.Collections.Generic;
using System.Text;

namespace PortSys.Tac.ClientServices.Kernel.Processing
{
    public class LaunchProcessParameters : ICommandParameters
    {
        /// <summary>
        /// Path to the working folder.
        /// </summary>
        /// <remarks>
        /// If this parameter is null, the process is started in the folder, where the executable file resides.
        /// </remarks>
        public string WorkingFolder { get; set; }
        
        /// <summary>
        /// Fully qualified file name. No expansion strings allowed.
        /// </summary>
        public string CommandLine { get; set; }

        /// <summary>
        /// Gets the FQFN of currently on the command line.
        /// </summary>
        public string CommandFqfn
        {
            get
            {
                string result = null;
                int i = 0;

                if (CommandLine.StartsWith("\""))
                {
                    i = CommandLine.IndexOf('"', 1);
                    if (i != -1)
                    {
                        result = CommandLine.Substring(0, i + 1);
                    }
                }
                else
                {
                    i = CommandLine.IndexOf(" ");
                    result = (i == -1 ? CommandLine : CommandLine.Substring(0, i));
                }

                return result;
            }
        }

        public string CommandArguments
        {
            get
            {
                string appName = CommandFqfn;
                return CommandLine.Substring(appName.Length);
            }
        }

        public bool AreValid()
        {
            return !string.IsNullOrEmpty(CommandLine);
        }
    }
}
