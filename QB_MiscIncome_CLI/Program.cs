using System;
using System.Collections.Generic;
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
                //// 1. Query the QuickBooks deposits
                //Console.WriteLine("\n1. Querying initial QuickBooks MiscIncome records...");
                //QueryMiscIncomes();

                //// 2. Add a single test record
                //Console.WriteLine("\n2. Adding a single test MiscIncome record...");
                //MiscIncome testRecord = CreateTestRecord();
                //AddSingleMiscIncome(testRecord);

                //// 3. Query again to show the added record
                //Console.WriteLine("\n3. Querying QuickBooks MiscIncome records after adding one record...");
                //QueryMiscIncomes();

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
                Memo = "TEST-001" // Use this field for CompanyID
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
        /// Queries and displays all MiscIncome records from QuickBooks
        /// </summary>
        private static void QueryMiscIncomes()
        {
            try
            {
                List<MiscIncome> incomes = MiscIncomeReader.QueryAllMiscIncomes();

                Console.WriteLine($"Found {incomes.Count} MiscIncome records in QuickBooks.");

                // Display the records
                foreach (var income in incomes)
                {
                    Console.WriteLine($"TxnID: {income.TxnID}");
                    Console.WriteLine($"Date: {income.DepositDate:yyyy-MM-dd}");
                    Console.WriteLine($"To Account: {income.DepositToAccount}");
                    Console.WriteLine($"Memo (CompanyID): {income.Memo}");
                    Console.WriteLine($"Total Amount: {income.TotalAmount:C}");
                    Console.WriteLine($"Line Count: {income.Lines.Count}");

                    // Show line details
                    foreach (var line in income.Lines)
                    {
                        Console.WriteLine($"  - {line.Amount:C} from {line.ReceivedFromName}, Account: {line.FromAccountName}, Memo: {line.Memo}");
                    }

                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error querying MiscIncome records: {ex.Message}");
                throw;
            }
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
                Console.WriteLine($"Error adding MiscIncome records: {ex.Message}");
                throw;
            }
        }
    }
}