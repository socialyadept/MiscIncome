using System;
using System.Collections.Generic;
using QB_MiscIncome_Lib;
using Serilog;

namespace QB_MiscIncome_CLI
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set up Serilog for logging
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File("Logs\\app.log", rollingInterval: RollingInterval.Day)
                .CreateLogger();

            try
            {
                // Query all Misc Income records from QuickBooks
                List<MiscIncome> deposits = MiscIncomeReader.QueryAllMiscIncomes();

                // Display the results
                if (deposits.Count > 0)
                {
                    Console.WriteLine("Retrieved the following deposits from QuickBooks:");
                    foreach (var deposit in deposits)
                    {
                        Console.WriteLine($"TxnID: {deposit.TxnID}, Date: {deposit.DepositDate}, Total Amount: {deposit.TotalAmount}, Account: {deposit.DepositToAccount}");
                        foreach (var line in deposit.Lines)
                        {
                            Console.WriteLine($"  - Customer: {line.ReceivedFromName}, Amount: {line.Amount}, From Account: {line.FromAccountName}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No deposit transactions found.");
                }
            }
            catch (Exception ex)
            {
                // Log any errors that occur during the query
                Log.Error(ex, "An error occurred while querying QuickBooks for deposit transactions.");
            }
            finally
            {
                // Ensure to close the logger cleanly at the end of execution
                Log.CloseAndFlush();
            }

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}