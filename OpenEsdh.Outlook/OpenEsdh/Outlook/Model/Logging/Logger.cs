namespace OpenEsdh.Outlook.Model.Logging
{
    using System;
    using System.Diagnostics;
    using System.Globalization;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;

    public class Logger
    {
        private static object _lock = new object();
        private static Logger _logger = null;
        private const int MaxEventLogEntryLength = 0x7530;
        private const string ServiceName = "OpenESDH.Outlook";
        private const string SourceName = "Application";

        private string EnsureLogMessageLimit(string logMessage)
        {
            if (logMessage.Length > 0x7530)
            {
                string str = string.Format(CultureInfo.CurrentCulture, "... | Log Message Truncated [ Limit: {0} ]", new object[] { 0x7530 });
                logMessage = logMessage.Substring(0, 0x7530 - str.Length);
                logMessage = string.Format(CultureInfo.CurrentCulture, "{0}{1}", new object[] { logMessage, str });
            }
            return logMessage;
        }

        private string GetSource()
        {
            if (!string.IsNullOrWhiteSpace(this.Source))
            {
                return this.Source;
            }
            try
            {
                Assembly entryAssembly = Assembly.GetEntryAssembly();
                if (entryAssembly == null)
                {
                    entryAssembly = Assembly.GetExecutingAssembly();
                }
                if (entryAssembly == null)
                {
                    entryAssembly = new StackTrace().GetFrames().Last<StackFrame>().GetMethod().Module.Assembly;
                }
                if (entryAssembly == null)
                {
                    return "Unknown";
                }
                return entryAssembly.GetName().Name;
            }
            catch
            {
                return "Unknown";
            }
        }

        private void Log(string message, EventLogEntryType entryType, string source)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(source))
                {
                    source = this.GetSource();
                }
                string str = this.EnsureLogMessageLimit(message);
                using (EventLog log = new EventLog("Application"))
                {
                    log.Source = "OpenESDH.Outlook";
                    log.BeginInit();
                    if (!EventLog.SourceExists(log.Source))
                    {
                        EventLog.CreateEventSource(log.Source, log.Log);
                    }
                    log.EndInit();
                    log.WriteEntry(source + ":" + message, entryType, 1);
                }
                if (Environment.UserInteractive)
                {
                    Console.WriteLine(message);
                }
            }
            catch
            {
            }
        }

        public void LogDebug(string message, bool debugLoggingEnabled, string source = "")
        {
            if (debugLoggingEnabled)
            {
                this.Log(message, EventLogEntryType.Information, source);
            }
        }

        public void LogException(Exception ex, string source = "")
        {
            if (ex == null)
            {
                throw new ArgumentNullException("ex");
            }
            if (Environment.UserInteractive)
            {
                Console.WriteLine(ex.ToString());
            }
            this.Log(ex.ToString(), EventLogEntryType.Error, source);
        }

        public void LogInformation(string message, string source = "")
        {
            this.Log(message, EventLogEntryType.Information, source);
        }

        public void LogWarning(string message, string source = "")
        {
            this.Log(message, EventLogEntryType.Warning, source);
        }

        public static Logger Current
        {
            get
            {
                if (_logger == null)
                {
                    lock (_lock)
                    {
                        if (_logger == null)
                        {
                            _logger = new Logger();
                        }
                    }
                }
                return _logger;
            }
        }

        public string Source { get; set; }
    }
}

