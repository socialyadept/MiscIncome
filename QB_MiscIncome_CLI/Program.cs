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
                // Present a menu of options to the user
                while (true)
                {
                    Console.Clear();
                    Console.WriteLine("QuickBooks Desktop MiscIncome Integration");
                    Console.WriteLine("=======================================");
                    Console.WriteLine("1. Query QuickBooks MiscIncome records");
                    Console.WriteLine("2. Add a single test MiscIncome record");
                    Console.WriteLine("3. Compare Excel records with QuickBooks");
                    Console.WriteLine("4. Add all MiscIncome records from Excel file");
                    Console.WriteLine("5. Exit");
                    Console.WriteLine();
                    Console.Write("Enter your choice (1-5): ");

                    string choice = Console.ReadLine();

                    switch (choice)
                    {
                        case "1":
                            Console.Clear();
                            Console.WriteLine("Querying QuickBooks MiscIncome records...");
                            QueryMiscIncomes();
                            Console.WriteLine("\nPress any key to return to the menu...");
                            Console.ReadKey();
                            break;

                        case "2":
                            Console.Clear();
                            Console.WriteLine("Adding a single test MiscIncome record...");
                            MiscIncome testRecord = CreateTestRecord();
                            AddSingleMiscIncome(testRecord);
                            Console.WriteLine("\nPress any key to return to the menu...");
                            Console.ReadKey();
                            break;

                        case "3":
                            Console.Clear();
                            Console.WriteLine("Comparing Excel records with QuickBooks records...");
                            if (File.Exists(excelFilePath))
                            {
                                CompareMiscIncomesFromExcel();
                            }
                            else
                            {
                                Console.WriteLine($"Error: Excel File not found at: {excelFilePath}");
                            }
                            Console.WriteLine("\nPress any key to return to the menu...");
                            Console.ReadKey();
                            break;

                        case "4":
                            Console.Clear();
                            Console.WriteLine("Adding all MiscIncome records from Excel file...");
                            if (File.Exists(excelFilePath))
                            {
                                AddAllMiscIncomesFromExcel();
                            }
                            else
                            {
                                Console.WriteLine($"Error: Excel File not found at: {excelFilePath}");
                            }
                            Console.WriteLine("\nPress any key to return to the menu...");
                            Console.ReadKey();
                            break;

                        case "5":
                            Console.WriteLine("Exiting application...");
                            return;

                        default:
                            Console.WriteLine("Invalid choice. Press any key to try again...");
                            Console.ReadKey();
                            break;
                    }
                }
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

                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
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
        /// Queries MiscIncome records from QuickBooks
        /// </summary>
        private static void QueryMiscIncomes()
        {
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
        /// Compares MiscIncome records from Excel with QuickBooks without adding them
        /// </summary>
        private static void CompareMiscIncomesFromExcel()
        {
            Console.WriteLine($"Parsing Excel file: {excelFilePath}");
            List<MiscIncome> excelIncomes = ExcelParser.ParseExcelFile(excelFilePath);

            if (excelIncomes.Count == 0)
            {
                Console.WriteLine("No records found in Excel file or error parsing file.");
                return;
            }

            Console.WriteLine($"Found {excelIncomes.Count} records in Excel file.");

            try
            {
                Console.WriteLine("\nComparing Excel records with QuickBooks records...");
                List<MiscIncome> newRecords = MiscIncomeComparator.CompareMiscIncomes(excelIncomes);

                // Ask user if they want to add new records
                if (newRecords.Count > 0)
                {
                    Console.WriteLine($"\nFound {newRecords.Count} new records that could be added to QuickBooks.");
                    Console.WriteLine("Would you like to add these new records to QuickBooks? (Y/N)");
                    string response = Console.ReadLine().Trim().ToUpper();

                    if (response == "Y")
                    {
                        Console.WriteLine($"\nAdding {newRecords.Count} new records to QuickBooks...");

                        try
                        {
                            using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                            {
                                var adder = new MiscIncomeAdder(qbSession);
                                adder.AddMiscIncomes(newRecords);
                                Console.WriteLine($"Successfully added {newRecords.Count} new MiscIncome records.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, "Error adding new MiscIncome records from comparison");
                            Console.WriteLine($"Error adding new MiscIncome records: {ex.Message}");
                            throw;
                        }
                    }
                    else
                    {
                        Console.WriteLine("No records will be added to QuickBooks.");
                    }
                }
                else
                {
                    Console.WriteLine("No new records to add to QuickBooks.");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error comparing MiscIncome records");
                Console.WriteLine($"Error comparing MiscIncome records: {ex.Message}");
                throw;
            }
        }
    }
}