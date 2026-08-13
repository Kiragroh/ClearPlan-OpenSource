using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.IO;
using System.Reflection;

namespace ClearPlan
{
    public static class Log
    {
        private static readonly LogFactory logFactory = new LogFactory();
        private static readonly Logger logger;

        static Log()
        {
            string logsDir = Path.Combine(GetAssemblyDirectory(), "Logs");

            // Stelle sicher, dass der Logs-Ordner existiert
            if (!Directory.Exists(logsDir))
            {
                Directory.CreateDirectory(logsDir);
            }

            string logFilePath = Path.Combine(logsDir, "globalLog.csv");

            var config = new LoggingConfiguration();

            // Allgemeiner Log (CSV mit Semikolon-Trennung)
            var logfile = new FileTarget("logfile")
            {
                FileName = logFilePath,
                Layout = "${date:format=yyyy-MM-dd HH\\:mm\\:ss};${level:uppercase=true};${message}${onexception:${newline}  ${exception:format=Message,StackTrace:separator=\r\n}}"
            };

            config.AddRule(LogLevel.Trace, LogLevel.Fatal, logfile);

            logFactory.Configuration = config;
            LogManager.ReconfigExistingLoggers();
            logger = logFactory.GetLogger("GeneralLogger");

            // Prüfe, ob Datei existiert, sonst Header schreiben
            if (!File.Exists(logFilePath))
            {
                File.AppendAllText(logFilePath, "Date;Level;Message;Exception\n");
            }
        }

        public static Logger GetLogger() => logger;

        public static Logger GetCustomLogger(string logFileName)
        {
            var config = new LoggingConfiguration();
            string logPath = Path.Combine(GetAssemblyDirectory(), "Logs", logFileName);

            var customLogFile = new FileTarget(logFileName)
            {
                FileName = logPath,
                Layout = "${longdate} ${level:uppercase=true} ${message} ${exception:format=ToString}"
            };

            config.AddRule(LogLevel.Trace, LogLevel.Fatal, customLogFile);

            var customFactory = new LogFactory();
            customFactory.Configuration = config;
            return customFactory.GetLogger(logFileName);
        }

        private static string GetAssemblyDirectory()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }
    }
    public class CustomLog
    {
        private readonly Logger logger;

        public CustomLog(string logFileName)
        {
            logger = Log.GetCustomLogger(logFileName);
        }

        public void Info(string message)
        {
            logger.Info(message);
        }

        public void Error(Exception ex, string message = "")
        {
            logger.Error(ex, message);
        }
    }
}