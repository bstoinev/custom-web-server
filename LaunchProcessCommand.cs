using Newtonsoft.Json;
using PortSys.Tac.ClientServices.Intelligence;
using PortSys.Tac.ClientServices.Kernel.Win32;
using System;

namespace PortSys.Tac.ClientServices.Kernel.Processing
{
    [CommandName("LaunchProcess")]
    public class LaunchProcessCommand : ICommand<LaunchProcessParameters, ICommandResult>
    {
        public LaunchProcessCommand()
        {
            Monitor = new CommandMonitor("EngineEventSource");
        }

        public CommandState State { get; set; }

        public LaunchProcessParameters Parameters { get; protected set; }

        public ICommandResult Result { get; protected set; }

        public IMonitor Monitor { get; set; }

        public void ReadParameters(string JsonString)
        {
            if (!string.IsNullOrEmpty(JsonString))
            {
                Parameters = JsonConvert.DeserializeObject<LaunchProcessParameters>(JsonString);
            }            
        }

        public void Run()
        {
            State = CommandState.Running;

            if (Parameters == null || !Parameters.AreValid())
            {
                throw new InvalidOperationException("Invalid parameters.");
            }

            string tf = Parameters.CommandFqfn;

            if (string.IsNullOrEmpty(tf))
            {
                throw new System.Data.SyntaxErrorException();
            }

            string ca = Parameters.CommandArguments;
            string wf = Parameters.WorkingFolder;

            using (var helper = new LaunchProcessCommandWin32Helper())
            {                
                helper.SpawnProcessToActiveConsole(tf, ca, wf);
            }

            State = CommandState.Completed;
        }
    }
}
