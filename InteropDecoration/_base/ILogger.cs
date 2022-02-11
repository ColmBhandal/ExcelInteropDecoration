using System.Runtime.InteropServices;

namespace InteropDecoration._base
{
    public interface ILogger
    {
        void Debug(string message);
        void Debug(string message, Exception e);
        void Info(string message);
        void Warn(string message, Exception ex);
        void Warn(string message);
        void Error(string message, Exception ex);
    }
}