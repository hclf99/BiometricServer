using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Xml;

namespace BiometricServer
{
    public enum LogLevel : short
    {
        Fatal,
        Error,
        Warning,
        Info,
        Debug,
        All
    }

    class LogData
    {
        public DateTime Timestamp;
        public LogLevel Level;
        public string Category;
        public string Message;
        public Exception Exception;
    }

    struct LogCategory
    {
        public string Name;
        public LogLevel Level;
    }

    public class Logger
    {
        //----------------------------------------------------------------
        public Logger(string identifier)
        //----------------------------------------------------------------
        {
            Directory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            Identifier = identifier;
            CurrentLogLevel = LogLevel.Info;

            WorkerThread = new Thread(new ParameterizedThreadStart(Worker));
            WorkerThread.Start(this);
        }

        //----------------------------------------------------------------
        public void Debug(string message, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {message}";
            Log(DateTime.Now, LogLevel.Debug, null, nMessage, null);
        }

        //----------------------------------------------------------------
        public void Info(string message, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {message}";
            Log(DateTime.Now, LogLevel.Info, null, nMessage, null);
        }

        //----------------------------------------------------------------
        public void Warning(string message, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {message}";
            Log(DateTime.Now, LogLevel.Warning, null, nMessage, null);
        }

        //----------------------------------------------------------------
        public void Error(string message, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {message}";
            Log(DateTime.Now, LogLevel.Error, null, nMessage, null);
        }

        //----------------------------------------------------------------
        public void Fatal(string message, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {message}";
            Log(DateTime.Now, LogLevel.Fatal, null, nMessage, null);
        }

        //----------------------------------------------------------------
        public void Debug(Exception exception, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {exception.Message}";
            Log(DateTime.Now, LogLevel.Debug, null, nMessage, exception);
        }

        //----------------------------------------------------------------
        public void Info(Exception exception, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {exception.Message}";
            Log(DateTime.Now, LogLevel.Info, null, nMessage, exception);
        }

        //----------------------------------------------------------------
        public void Warning(Exception exception, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {exception.Message}";
            Log(DateTime.Now, LogLevel.Warning, null, nMessage, exception);
        }

        //----------------------------------------------------------------
        public void Error(Exception exception, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {exception.Message}";
            Log(DateTime.Now, LogLevel.Error, null, nMessage, exception);
        }

        //----------------------------------------------------------------
        public void Fatal(Exception exception, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {exception.Message}";
            Log(DateTime.Now, LogLevel.Fatal, null, nMessage, exception);
        }

        //----------------------------------------------------------------
        public void Debug(string category, string message, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {message}";
            Log(DateTime.Now, LogLevel.Debug, category, nMessage, null);
        }

        //----------------------------------------------------------------
        public void Info(string category, string message, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {message}";
            Log(DateTime.Now, LogLevel.Info, category, nMessage, null);
        }

        //----------------------------------------------------------------
        public void Warning(string category, string message, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {message}";
            Log(DateTime.Now, LogLevel.Warning, category, nMessage, null);
        }

        //----------------------------------------------------------------
        public void Error(string category, string message, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {message}";
            Log(DateTime.Now, LogLevel.Error, category, nMessage, null);
        }

        //----------------------------------------------------------------
        public void Fatal(string category, string message, [CallerFilePath] string file = null, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        //----------------------------------------------------------------
        {
            string nMessage = $"{file.Split('\\').Last()} - {caller} - {lineNumber}: {message}";
            Log(DateTime.Now, LogLevel.Fatal, category, nMessage, null);
        }

        //----------------------------------------------------------------
        public void Log(DateTime timestamp, LogLevel level, string category, string message, Exception exception)
        //----------------------------------------------------------------
        {
            if (IsStopped)
            {
                return;
            }

            LogData logData = new LogData();

            logData.Timestamp = timestamp;
            logData.Level = level;
            logData.Category = category;
            logData.Message = message;
            logData.Exception = exception;

            lock (this)
            {
                LogDataQueue.Enqueue(logData);
            }
        }

        //----------------------------------------------------------------
        public void StopAndWait()
        //----------------------------------------------------------------
        {
            IsStopped = true;

            WorkerThread.Join();
        }

        //----------------------------------------------------------------
        public bool IsDebugEnabled
        //----------------------------------------------------------------
        {
            get
            {
                return LogLevel.Debug <= CurrentLogLevel;
            }
        }

        //----------------------------------------------------------------
        public bool IsInfoEnabled
        //----------------------------------------------------------------
        {
            get
            {
                return LogLevel.Info <= CurrentLogLevel;
            }
        }

        //----------------------------------------------------------------
        public bool IsWarningEnabled
        //----------------------------------------------------------------
        {
            get
            {
                return LogLevel.Warning <= CurrentLogLevel;
            }
        }

        //----------------------------------------------------------------
        public bool IsErrorEnabled
        //----------------------------------------------------------------
        {
            get
            {
                return LogLevel.Error <= CurrentLogLevel;
            }
        }

        //----------------------------------------------------------------
        public bool IsFatalEnabled
        //----------------------------------------------------------------
        {
            get
            {
                return LogLevel.Fatal <= CurrentLogLevel;
            }
        }

        //----------------------------------------------------------------
        public string Directory { get; set; }
        //----------------------------------------------------------------

        //----------------------------------------------------------------
        public string Identifier { get; private set; }
        //----------------------------------------------------------------

        //----------------------------------------------------------------
        public string LogFileName
        //----------------------------------------------------------------
        {
            get
            {
                return $"{Directory}\\{Identifier}_{DateTime.Today.ToString("yyyyMMdd")}.log";
            }
        }

        //----------------------------------------------------------------
        public string CfgFileName
        //----------------------------------------------------------------
        {
            get
            {
                return $"L:\\cfg\\{Identifier}.cfg";
            }
        }

        //----------------------------------------------------------------
        public LogLevel CurrentLogLevel
        //----------------------------------------------------------------
        {
            get;
            private set;
        }

        //----------------------------------------------------------------
        private void Worker(object parameter)
        //----------------------------------------------------------------
        {
            Logger thisLogger = (Logger)parameter;
            LogData logData = null;

            while (!(thisLogger.IsStopped && LogDataQueue.Count == 0))
            {
                if (LogDataQueue.Count == 0)
                {
                    Thread.Sleep(10);
                    continue;
                }

                try
                {
                    ReadConfiguration();
                }
                catch (Exception e)
                {
                    System.Diagnostics.Debug.WriteLine("Read configuration file <" + CfgFileName + "> failed!");
                    System.Diagnostics.Debug.WriteLine(e.Message);
                }
                finally
                {
                    lock (this)
                    {
                        logData = LogDataQueue.Dequeue();
                    }

                    if (null != logData)
                    {
                        try
                        {
                            string logText = LogDataToString(logData);

                            if (null != logText)
                            {
                                System.Diagnostics.Debug.WriteLine(logText);

                                File.AppendAllText(LogFileName, logText + "\r\n");
                            }
                        }
                        catch (Exception e)
                        {
                            System.Diagnostics.Debug.WriteLine("Write to logfile <" + LogFileName + "> failed!");
                            System.Diagnostics.Debug.WriteLine(e.Message);
                        }
                    }
                }
            }
        }

        //----------------------------------------------------------------
        private void ReadConfiguration()
        //----------------------------------------------------------------
        {
            Categories.Clear();

            if (!File.Exists(CfgFileName))
            {
                return;
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(CfgFileName);

            XmlNode levelNode = xmlDoc.SelectSingleNode("config/level");

            if (null != levelNode)
            {
                try
                {
                    CurrentLogLevel = (LogLevel)Enum.Parse(typeof(LogLevel), levelNode.InnerText, true);
                }
                catch (Exception e)
                {
                    // Ignore exceptions
                    System.Diagnostics.Debug.WriteLine(e.Message);
                }
            }

            xmlDoc.SelectNodes("config/category").OfType<XmlNode>().ToList().ForEach(n =>
            {
                XmlNode categoryLevelNode = n.Attributes.GetNamedItem("level");
                LogLevel categoryLevel = CurrentLogLevel;

                if (null != categoryLevelNode)
                {
                    try
                    {
                        categoryLevel = (LogLevel)Enum.Parse(typeof(LogLevel), categoryLevelNode.Value, true);
                    }
                    catch (Exception e)
                    {
                        // Ignore exceptions
                        System.Diagnostics.Debug.WriteLine(e.Message);
                    }
                }

                LogCategory category;
                category.Name = n.InnerText;
                category.Level = categoryLevel;

                Categories.Add(category);
            });
        }

        //----------------------------------------------------------------
        private string LogDataToString(LogData logData)
        //----------------------------------------------------------------
        {
            if (logData.Category != null && logData.Category.Length > 0)
            {
                bool categoryEnabled = false;

                foreach (LogCategory category in Categories)
                {
                    if (logData.Level <= category.Level)
                    {
                        if (logData.Category.IndexOf(category.Name) == 0)
                        {
                            categoryEnabled = true;
                            break;
                        }
                    }
                }

                if (!categoryEnabled)
                {
                    return null;
                }
            }
            else
            {
                if (CurrentLogLevel < logData.Level)
                {
                    // Only send messages that have a lower or equal level as the current level
                    return null;
                }
            }

            string time = logData.Timestamp.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string levelText = logData.Level.ToString().ToUpper();

            string logMessage = time + " ";
            logMessage += "[" + levelText + "] ";

            if (logData.Category != null && logData.Category.Length > 0)
            {
                logMessage += "[" + logData.Category + "] ";
            }

            logMessage += logData.Message;

            if (null != logData.Exception)
            {
                logMessage += "\r\n";
                logMessage += logData.Exception.Source + "\r\n";
                logMessage += logData.Exception.StackTrace;
            }

            return logMessage;
        }

        private Thread WorkerThread;
        private Queue<LogData> LogDataQueue = new Queue<LogData>();
        private List<LogCategory> Categories = new List<LogCategory>();
        public bool IsStopped { get; private set; } = false;
    }

    public class LogManager
    {
        //----------------------------------------------------------------
        public static Logger DefaultLogger
        //----------------------------------------------------------------
        {
            get
            {
                return GetLogger("Default");
            }
        }

        //----------------------------------------------------------------
        public static Logger AssistLogger
        //----------------------------------------------------------------
        {
            get
            {
                return GetLogger("Assist");
            }
        }

        //----------------------------------------------------------------
        public static Logger GetLogger(string identifier)
        //----------------------------------------------------------------
        {
            lock (LockObj)
            {
                if (Loggers.ContainsKey(identifier))
                {
                    return Loggers[identifier];
                }
                else
                {
                    Logger logger = new Logger(identifier);
                    Loggers.Add(identifier, logger);

                    return logger;
                }
            }
        }

        //----------------------------------------------------------------
        public static void StopAllAndWait()
        //----------------------------------------------------------------
        {
            foreach (KeyValuePair<string, Logger> pair in Loggers)
            {
                pair.Value.StopAndWait();
            }

            Loggers.Clear();
        }

        private static Dictionary<string, Logger> Loggers = new Dictionary<string, Logger>();
        private static object LockObj = new object();
    }
}
