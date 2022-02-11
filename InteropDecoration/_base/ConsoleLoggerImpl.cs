using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace InteropDecoration._base
{
    internal class ConsoleLoggerImpl : ILogger
    {
        public void Debug(string message)
        {
            LogMessage(message, "DEBUG");
        }

        public void Debug(string message, Exception ex)
        {
            LogMessage(message, ex, "DEBUG");
        }

        public void Info(string message)
        {
            LogMessage(message, "INFO");
        }

        public void Warn(string message)
        {
            LogMessage(message, "WARN");
        }

        public void Warn(string message, Exception ex)
        {
            LogMessage(message, ex, "WARN");
        }

        public void Error(string message, Exception ex)
        {
            LogMessage(message, "ERROR");
        }

        private void LogMessage(string message, Exception ex, string logLevelPrefix)
        {
            StringBuilder sb = new StringBuilder(message)
                .AppendLine("Included exception: " + ex.StackTrace);
            string mesageWithException = sb.ToString();
            LogMessage(logLevelPrefix, mesageWithException);
        }

        private void LogMessage(string message, string logLevelPrefix)
        {
            Console.WriteLine($"{logLevelPrefix}: {message}");
        }
    }
}
