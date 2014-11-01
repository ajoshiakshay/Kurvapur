using System;
using System.Collections.Generic;
using System.Text;
using log4net;
using log4net.Config;

namespace KurvapurSevaApp.Utilities
{
    class Logger
    {
        private static readonly ILog logger = LogManager.GetLogger("");

        static Logger()
        {
            DOMConfigurator.Configure();
        }

        public void Log(string strLevel, string strMsg)
        {
            switch (strLevel)
            {
                case "DEBUG" :
                    logger.Debug(strMsg);
                    break;
                case "INFO":
                    logger.Info(strMsg);
                    break;
                case "WARN":
                    logger.Warn(strMsg);
                    break;
                case "ERROR":
                    logger.Error(strMsg);
                    break;
                default:
                    logger.Fatal(strMsg);
                    break;
            }
        }
    }
}
