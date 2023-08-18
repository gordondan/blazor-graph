using Serilog;

namespace BlazorGraph
{
    public static class LoggerExtensions
    {
        public static void Debug(this ILogger logger, string message, bool includeTimestamp)
        {
            if (includeTimestamp)
            {
                message = $"{DateTime.UtcNow:yyyy-MM-dd HH:mm:ss.fff zzz} - {message}";
            }

            logger.Debug(message);
        }
    }
}