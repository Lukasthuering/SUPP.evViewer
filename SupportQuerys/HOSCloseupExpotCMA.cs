using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BiCA.Sabas.Extension.V2;
using BiCA.Sabas.Support.V2;
using BiCA.Sabas.Support.V2.Core;

namespace SupportQuerys
{
    class HOSCloseupExpotCMA : ManagedScript
    {
        Log _log;
        IComputer sim;

        protected override void OnErrorOccurred(Exception exception)
        {
            _log.Info(exception.Message);
            _log.Info(@"Ended with error see logs C:\Temp\read_file_log.txt !");
        }

        protected override void OnExecuting()
        {
            
        }

        protected override void OnFinalizing()
        {
            _log.Info("Skript finsihed !");
        }

        protected override void OnInitializing()
        {
            _log = new Log(Runtime.System, "time_log.txt", @"C:\Temp\", false);
            sim = Runtime.System.SiteServer;
            _log.Info("Skript gestartet !");
        }
        
    }
}

