using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using log4net.Config;

namespace MSExcelToMongoConnector
{
    public class Logger
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Logger));

        public static void initialize()
        {
            BasicConfigurator.Configure();
        }
        public static void logMessage(string message, LogLevel level)
        {
            
            if (level == LogLevel.ERROR)
            {
                log.Error(message);
            } else if (level == LogLevel.WARN)
            {
                log.Warn(message);
            } else
            {
                log.Info(message);
            }
        }
        public enum LogLevel
        {
            INFO,
            WARN,
            ERROR
        }
    }

   
}
