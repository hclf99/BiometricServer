using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Topshelf;

namespace BiometricServer
{
    public class WinService
    {
        //----------------------------------------------------------------
        public void Start()
        //----------------------------------------------------------------
        {

        }

        //----------------------------------------------------------------
        public void Stop()
        //----------------------------------------------------------------
        {

        }

        //----------------------------------------------------------------
        public void Shutdown()
        //----------------------------------------------------------------
        {

        }

        //----------------------------------------------------------------
        private String ReadSetting(string key)
        //----------------------------------------------------------------
        {
            var appSettings = ConfigurationManager.AppSettings;

            try
            {
                string result = appSettings[key] ?? "Not Found";
                LogManager.DefaultLogger.Info("Parameter read: " + key + " -> " + result);
                return result;
            }
            catch (ConfigurationErrorsException)
            {
                LogManager.DefaultLogger.Error("Error reading app setting: " + key);
                return "nil";
            }
        }
    }

    //----------------------------------------------------------------
    internal static class ConfigureService
    //----------------------------------------------------------------
    {
        internal static void Configure()
        {
            HostFactory.Run(configure =>
            {
                configure.Service<WinService>(service =>
                {
                    service.ConstructUsing(s => new WinService());
                    service.WhenStarted(s => s.Start());
                    service.WhenStopped(s => s.Stop());
                    service.WhenShutdown(s => s.Shutdown());
                });

                configure.RunAsLocalSystem();
                configure.StartAutomatically();
                configure.EnableShutdown();
                configure.SetServiceName("Biometric");
                configure.SetDisplayName("Biometric");
                configure.SetDescription("Service for Biometric machine");
            });
        }
    }

}
