using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using log4net.Appender;
using log4net.Config;

namespace BonenLawyer
{
    public static class Log
    {
        private static ILog _logger;

        //internal for testing...
        internal static ILog Logger { set { _logger = value; } }

        private static string _filePath;
        private const string DefaultConversionPatern = "%date %-5level- %message%newline";

        public static bool LoggerInitialized { get { return _logger != null; } }

        //configure log4Net via code instead of xml
        public static void InitLog(string completeLogFile)
        {

            _filePath = completeLogFile;

            var hierarchy = (log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository();
            hierarchy.Root.RemoveAllAppenders();

            var pl = new log4net.Layout.PatternLayout { ConversionPattern = DefaultConversionPatern };
            pl.ActivateOptions();
            var fileAppender = new RollingFileAppender()
            {
                AppendToFile = true,
                LockingModel = new FileAppender.MinimalLock(),
                File = _filePath,
                Layout = pl,
            };

            fileAppender.MaxFileSize = 10 * 1024 * 1024;
            fileAppender.RollingStyle = RollingFileAppender.RollingMode.Size;
            fileAppender.MaxSizeRollBackups = 5;

            fileAppender.ActivateOptions();

            BasicConfigurator.Configure(fileAppender);

            _logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        }

        public static void Exception(Exception ex)
        {
            if (_logger == null) return; //we want to use Log in code under test without having to bind the logger....
            _logger.Error(ex.ToString());
        }

        public static void Error(String error, params object[] arg)
        {
            if (_logger == null) return; //we want to use Log in code under test without having to bind the logger....
            _logger.ErrorFormat(error, arg);
        }


        public static void Info(String traceMessage, params object[] arg)
        {
            if (_logger == null) return; //we want to use Log in code under test without having to bind the logger....
            _logger.InfoFormat(traceMessage, arg);
        }


        public static void Debug(String traceMessage, params object[] arg)
        {
            if (_logger == null) return; //we want to use Log in code under test without having to bind the logger....
            _logger.DebugFormat(traceMessage, arg);
        }
    }
}
