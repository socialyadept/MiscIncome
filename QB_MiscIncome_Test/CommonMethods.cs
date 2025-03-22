using QB_MiscIncome_CLI;
using Serilog;

namespace QB_MiscIncome_Test
{
    public class CommonMethods
    {
        public static void EnsureLogFileClosed()
        {
            Log.CloseAndFlush(); // Forces Serilog to release the log file
            Thread.Sleep(1000); // Allow time for Serilog to fully release the file handle
        }

        public static void DeleteOldLogFiles()
        {
            string logDirectory = "logs";
            string logPattern = "qb_sync*.log";

            if (Directory.Exists(logDirectory))
            {
                foreach (var logFile in Directory.GetFiles(logDirectory, logPattern, SearchOption.TopDirectoryOnly))
                {
                    TryDeleteFile(logFile);
                }
            }
        }

        public static void EnsureLogFileExists(string logFile)
        {
            const int MAX_RETRIES = 10;
            int delayMilliseconds = 200;

            for (int attempt = 0; attempt < MAX_RETRIES; attempt++)
            {
                if (File.Exists(logFile))
                {
                    return; // Log file found, continue test
                }

                Thread.Sleep(delayMilliseconds); // Wait and retry
            }

            throw new FileNotFoundException($"Log file '{logFile}' was not created after test execution.");
        }

        public static void TryDeleteFile(string filePath)
        {
            const int maxRetries = 10;
            int delayMilliseconds = 200;

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    File.Delete(filePath);
                    return;
                }
                catch (IOException)
                {
                    Thread.Sleep(delayMilliseconds);
                    delayMilliseconds += 200;
                }
            }
        }

        public static string GetLatestLogFile()
        {
            string logDirectory = "logs";
            string logPattern = "qb_sync*.log";

            string[] logFiles = Directory.GetFiles(logDirectory, logPattern, SearchOption.TopDirectoryOnly);
            if (logFiles.Length == 0)
            {
                throw new FileNotFoundException("No log files found after test run.");
            }

            return logFiles.OrderByDescending(File.GetLastWriteTimeUtc).First();
        }

        public static void ResetLogger()
        {
            LoggerConfig.ResetLogger();
            LoggerConfig.ConfigureLogging();
        }
    }
}
