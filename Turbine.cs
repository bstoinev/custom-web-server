using PortSys.Tac.ClientServices.Hosting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace PortSys.Tac.ClientServices.Kernel
{
    public class Turbine : MarshalByRefObject, ITcsKernel
    {
        public Turbine()
        {
            Monitor = new CommandMonitor("EngineEventSource");
        }

        public CommandMonitor Monitor { get; private set; }

        public HttpServer Server { get; private set; }

        #region MarshalByRefObject support

        public override object InitializeLifetimeService()
        {
            return null;
        }

        #endregion

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Monitor.Error("Unhandled exception occurred:\n{0}", e.ExceptionObject.ToString());            
            Dispose();
        }

        private void CommandInterpreter_RebootKernelRequested()
        {
            if (RebootDelegate == null)
            {
                Monitor.Warning("Unable to reboot kernel. Reboot handler not set.");
            }
            else
            {
                ThreadPool.QueueUserWorkItem(new WaitCallback((state) => {
                    CommandInterpreter.ShutDown();
                    Server.Shutdown();
                    RebootDelegate();
                }));
            }
        }

        public void Run()
        {
            Server = new HttpServer();
            CommandInterpreter.RebootKernelRequested += CommandInterpreter_RebootKernelRequested;
            Monitor.Debug("Client Services engine started successfully.\nKernel loaded from: {0}", System.Reflection.Assembly.GetExecutingAssembly().GetName());
        }

        public void Dispose()
        {
            CommandInterpreter.RebootKernelRequested -= CommandInterpreter_RebootKernelRequested;

            if (Server != null)
            {
                Server.Dispose();
            }

            Monitor.Debug("Client Services engine stopped.");
#if DEBUG
            Monitor.Debug("*******************************");
#endif
        }

        public Action RebootDelegate { get; set; }
    }
}
