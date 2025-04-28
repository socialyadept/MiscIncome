using System;
using System.Collections.Generic;
using System.IO;
using QB_MiscIncome_Lib;
using Serilog;

namespace QB_MiscIncome_CLI
{
    class Program
    {
        // Update the file path to point to your Excel file
        private static readonly string excelFilePath = GetExcelFilePath();

        private static string GetExcelFilePath()
        {
            // Get the current directory (e.g., bin\Debug)
            string currentDir = Environment.CurrentDirectory;

            // Navigate up two levels to reach the project folder
            string projectDir = Directory.GetParent(Directory.GetParent(Directory.GetParent(currentDir).FullName).FullName).FullName;

            // Combine the project directory with the file name
            string computedFilePath = Path.Combine(projectDir, "credit-nonvendor.xlsx");
            Console.WriteLine("Computed file path: " + computedFilePath);
            return computedFilePath;
        }

        static void Main(string[] args)
        {
            // Set up Serilog for logging
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File("Logs\\app.log", rollingInterval: RollingInterval.Day)
                .CreateLogger();

            Console.WriteLine("QuickBooks Desktop MiscIncome Integration");
            Console.WriteLine("=======================================");

            try
            {
                // 1. Query the QuickBooks deposits
                Console.WriteLine("\n1. Querying initial QuickBooks MiscIncome records...");
                QueryMiscIncomes();

                // 2. Add a single test record
                Console.WriteLine("\n2. Adding a single test MiscIncome record...");
                MiscIncome testRecord = CreateTestRecord();
                AddSingleMiscIncome(testRecord);

                // 3. Query again to show the added record
                Console.WriteLine("\n3. Querying QuickBooks MiscIncome records after adding one record...");
                QueryMiscIncomes();

                // 4. Add all records from Excel file
                Console.WriteLine("\n4. Adding all MiscIncome records from Excel file...");
                if (File.Exists(excelFilePath))
                {
                    AddAllMiscIncomesFromExcel();
                }
                else
                {
                    throw new Exception($"Excel File not found at: {excelFilePath}");
                }

                // 5. Query all the results
                Console.WriteLine("\n5. Querying all QuickBooks MiscIncome records after adding Excel records...");
                QueryMiscIncomes();

            }
            catch (Exception ex)
            {
                Log.Error(ex, "An error occurred in the application");
                Console.WriteLine($"An error occurred: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                }
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        /// <summary>
        /// Creates a sample test record for demonstration
        /// </summary>
        /// <returns>A sample MiscIncome record</returns>
        private static MiscIncome CreateTestRecord()
        {
            MiscIncome testIncome = new MiscIncome
            {
                DepositDate = DateTime.Now,
                DepositToAccount = "Checking",
                Memo = "TEST-001" // Use this field for Child ID
            };

            // Add a single line
            testIncome.Lines.Add(new MiscIncomeLine
            {
                ReceivedFromName = "Misc Income",
                FromAccountName = "Sales",
                Amount = 1000.00,
                Memo = "Test deposit"
            });

            return testIncome;
        }

        /// <summary>
        /// Queries and prints all MiscIncome records from QuickBooks in a formatted table
        /// </summary>
        private static void PrintAllMiscIncomes()
        {
            try
            {
                List<MiscIncome> incomes = MiscIncomeReader.QueryAllMiscIncomes();

                if (incomes.Count == 0)
                {
                    Console.WriteLine("No MiscIncome records found in QuickBooks.");
                    return;
                }

                // Calculate total of all deposits
                double grandTotal = 0;
                int totalLines = 0;

                // Print header
                Console.WriteLine("\n============================================= MISC INCOME LIST =============================================");
                Console.WriteLine(String.Format("{0,-10} {1,-12} {2,-25} {3,-15} {4,-25} {5,-15}",
                    "Date", "Child ID", "Deposit Account", "Total Amount", "Line Count", "TxnID"));
                Console.WriteLine("--------------------------------------------------------------------------------------------------------");

                // Print each MiscIncome
                foreach (var income in incomes)
                {
                    Console.WriteLine(String.Format("{0,-10:d} {1,-12} {2,-25} {3,-15:C} {4,-25} {5,-15}",
                        income.DepositDate,
                        TruncateString(income.Memo, 12),
                        TruncateString(income.DepositToAccount, 25),
                        income.TotalAmount,
                        income.Lines.Count,
                        TruncateString(income.TxnID, 15)));

                    grandTotal += income.TotalAmount;
                    totalLines += income.Lines.Count;

                    // Print lines for this deposit
                    if (income.Lines.Count > 0)
                    {
                        // Print line header
                        Console.WriteLine(String.Format("   {0,-25} {1,-25} {2,-15} {3,-30}",
                            "Account", "Received From", "Amount", "Memo"));
                        Console.WriteLine("   -------------------------------------------------------------------------------------");

                        // Print each line
                        foreach (var line in income.Lines)
                        {
                            Console.WriteLine(String.Format("   {0,-25} {1,-25} {2,-15:C} {3,-30}",
                                TruncateString(line.FromAccountName, 25),
                                TruncateString(line.ReceivedFromName, 25),
                                line.Amount,
                                TruncateString(line.Memo, 30)));
                        }
                        Console.WriteLine(); // Add space between deposits
                    }
                }

                // Print footer with totals
                Console.WriteLine("--------------------------------------------------------------------------------------------------------");
                Console.WriteLine($"SUMMARY: {incomes.Count} deposits, {totalLines} deposit lines     Total Amount: {grandTotal:C}");
                Console.WriteLine("========================================================================================================");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error querying and displaying MiscIncome records");
                Console.WriteLine($"Error querying MiscIncome records: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Now we can remove this method since we've moved it to MiscIncomeReader
        /// 
        /// If you're updating the program, you should delete this method
        /// </summary>
        private static void QueryMiscIncomes()
        {
            // This method should be removed - functionality has been moved to MiscIncomeReader.QueryMiscIncomes()
            MiscIncomeReader.QueryMiscIncomes();
        }

        /// <summary>
        /// Adds a single MiscIncome record to QuickBooks
        /// </summary>
        /// <param name="income">The MiscIncome record to add</param>
        private static void AddSingleMiscIncome(MiscIncome income)
        {
            try
            {
                using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    var adder = new MiscIncomeAdder(qbSession);
                    adder.AddMiscIncomes(new List<MiscIncome> { income });
                    Console.WriteLine("Successfully added test MiscIncome record.");
                    Console.WriteLine($"TxnID assigned: {income.TxnID}");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to add test MiscIncome record");
                Console.WriteLine($"Failed to add test MiscIncome record: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Adds all MiscIncome records from the Excel file to QuickBooks
        /// </summary>
        private static void AddAllMiscIncomesFromExcel()
        {
            Console.WriteLine($"Parsing Excel file: {excelFilePath}");
            List<MiscIncome> incomes = ExcelParser.ParseExcelFile(excelFilePath);

            if (incomes.Count == 0)
            {
                Console.WriteLine("No records found in Excel file or error parsing file.");
                return;
            }

            Console.WriteLine($"Found {incomes.Count} records in Excel file.");

            try
            {
                using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    var adder = new MiscIncomeAdder(qbSession);
                    adder.AddMiscIncomes(incomes);
                    Console.WriteLine($"Successfully added {incomes.Count} MiscIncome records.");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error adding MiscIncome records from Excel");
                Console.WriteLine($"Error adding MiscIncome records: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Helper function to truncate strings for display formatting
        /// </summary>
        private static string TruncateString(string input, int maxLength)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            return input.Length <= maxLength ? input : input.Substring(0, maxLength - 3) + "...";
        }
    }
}