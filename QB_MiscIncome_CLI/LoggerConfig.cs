using Serilog;

namespace QB_MiscIncome_CLI
{
    public static class LoggerConfig
    {
        public static void ConfigureLogging()
        {
            // Always close and flush any existing logger
            Log.CloseAndFlush();

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .WriteTo.File("logs/qb_sync.log",
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 7,
                    fileSizeLimitBytes: 5_000_000,
                    rollOnFileSizeLimit: true,
                    restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Information)
                .CreateLogger();
        }

        public static void ResetLogger()
        {
            // Optional: This method just ensures Serilog is fully closed.
            // You can call it if you need to forcibly clear logs between uses.
            Log.CloseAndFlush();
        }
    }


}
