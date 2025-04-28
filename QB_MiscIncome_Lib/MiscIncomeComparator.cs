using System;
using System.Collections.Generic;
using System.Linq;
using Serilog;

namespace QB_MiscIncome_Lib
{
    /// <summary>
    /// Class for comparing MiscIncome records between Excel and QuickBooks
    /// </summary>
    public class MiscIncomeComparator
    {
        // Define record status constants to use internally since we can't add a property to MiscIncome
        private const string STATUS_UNCHANGED = "Unchanged";
        private const string STATUS_DIFFERENT = "Different";
        private const string STATUS_NEW = "New";
        private const string STATUS_MISSING = "Missing";

        /// <summary>
        /// Compares MiscIncome records from Excel with those in QuickBooks
        /// </summary>
        /// <param name="excelIncomes">List of MiscIncome records from Excel</param>
        /// <returns>List of MiscIncome records with comparison results</returns>
        public static List<MiscIncome> CompareMiscIncomes(List<MiscIncome> excelIncomes)
        {
            Log.Information("MiscIncomeComparator Initialized.");

            // Read QuickBooks MiscIncome records
            List<MiscIncome> qbIncomes = MiscIncomeReader.QueryAllMiscIncomes();

            // Convert QuickBooks and Excel records into dictionaries for quick lookup
            // Using Memo field (Child ID) as the key
            var qbIncomesDict = qbIncomes.ToDictionary(i => i.Memo, i => i);
            var excelIncomesDict = excelIncomes.ToDictionary(i => i.Memo, i => i);

            // Track records by status using separate dictionaries since we can't add a status property
            Dictionary<string, MiscIncome> newRecords = new Dictionary<string, MiscIncome>();
            Dictionary<string, MiscIncome> differentRecords = new Dictionary<string, MiscIncome>();
            Dictionary<string, MiscIncome> unchangedRecords = new Dictionary<string, MiscIncome>();
            Dictionary<string, MiscIncome> missingRecords = new Dictionary<string, MiscIncome>();

            // Print header for comparison results
            Console.WriteLine("\n==================== COMPARISON RESULTS ====================");

            // Iterate through Excel records to compare with QB records
            foreach (var excelIncome in excelIncomes)
            {
                if (qbIncomesDict.TryGetValue(excelIncome.Memo, out var qbIncome))
                {
                    // Record exists in both, compare details
                    if (AreMiscIncomesEqual(qbIncome, excelIncome))
                    {
                        Console.WriteLine($"UNCHANGED: MiscIncome record with Memo '{excelIncome.Memo}'");
                        Log.Information($"MiscIncome record with Memo '{excelIncome.Memo}' is Unchanged.");
                        unchangedRecords[excelIncome.Memo] = excelIncome;
                    }
                    else
                    {
                        Console.WriteLine($"DIFFERENT: MiscIncome record with Memo '{excelIncome.Memo}' needs updating in QuickBooks.");
                        Log.Information($"MiscIncome record with Memo '{excelIncome.Memo}' is Different - needs updating in QuickBooks.");
                        differentRecords[excelIncome.Memo] = excelIncome;

                        // Print differences
                        PrintMiscIncomeDifferences(qbIncome, excelIncome);
                    }
                }
                else
                {
                    // Record does not exist in QB, mark as new
                    Console.WriteLine($"NEW: MiscIncome record with Memo '{excelIncome.Memo}' not found in QuickBooks.");
                    Log.Information($"MiscIncome record with Memo '{excelIncome.Memo}' is New.");
                    newRecords[excelIncome.Memo] = excelIncome;
                }
            }

            // Check for records that exist in QB but not in the Excel file
            foreach (var qbIncome in qbIncomes)
            {
                if (!excelIncomesDict.ContainsKey(qbIncome.Memo))
                {
                    Console.WriteLine($"MISSING: MiscIncome record with Memo '{qbIncome.Memo}' exists in QuickBooks but is Missing from Excel file.");
                    Log.Information($"MiscIncome record with Memo '{qbIncome.Memo}' exists in QuickBooks but is Missing from Excel file.");
                    missingRecords[qbIncome.Memo] = qbIncome;
                }
            }

            // Print summary statistics
            Console.WriteLine("\n==================== COMPARISON SUMMARY ====================");
            Console.WriteLine($"Total records in QuickBooks: {qbIncomes.Count}");
            Console.WriteLine($"Total records in Excel: {excelIncomes.Count}");
            Console.WriteLine($"Unchanged records: {unchangedRecords.Count}");
            Console.WriteLine($"Different records (need update): {differentRecords.Count}");
            Console.WriteLine($"New records (not in QuickBooks): {newRecords.Count}");
            Console.WriteLine($"Missing records (not in Excel): {missingRecords.Count}");
            Console.WriteLine("==========================================================");

            // Merge all records for a comprehensive list
            List<MiscIncome> allRecords = new List<MiscIncome>();
            allRecords.AddRange(unchangedRecords.Values);
            allRecords.AddRange(differentRecords.Values);
            allRecords.AddRange(newRecords.Values);
            allRecords.AddRange(missingRecords.Values);

            Log.Information("MiscIncomeComparator Completed");

            // Return the combined list with records that need to be added
            return newRecords.Values.ToList();
        }

        /// <summary>
        /// Compares two MiscIncome records to determine if they are equal
        /// </summary>
        /// <param name="qbIncome">MiscIncome record from QuickBooks</param>
        /// <param name="excelIncome">MiscIncome record from Excel</param>
        /// <returns>True if records are effectively equal, false otherwise</returns>
        private static bool AreMiscIncomesEqual(MiscIncome qbIncome, MiscIncome excelIncome)
        {
            // Check basic properties
            if (qbIncome.DepositToAccount != excelIncome.DepositToAccount)
                return false;

            // Check line count
            if (qbIncome.Lines.Count != excelIncome.Lines.Count)
                return false;

            // Check if total amounts are significantly different (using tolerance for floating point comparison)
            const double tolerance = 0.01; // 1 cent tolerance
            if (Math.Abs(qbIncome.TotalAmount - excelIncome.TotalAmount) > tolerance)
                return false;

            // Create lookup dictionary for lines in QB record
            Dictionary<string, List<MiscIncomeLine>> qbLinesByAccount = new Dictionary<string, List<MiscIncomeLine>>();

            foreach (var line in qbIncome.Lines)
            {
                string key = line.FromAccountName;

                if (!qbLinesByAccount.ContainsKey(key))
                    qbLinesByAccount[key] = new List<MiscIncomeLine>();

                qbLinesByAccount[key].Add(line);
            }

            // Compare each Excel line with QB lines
            foreach (var excelLine in excelIncome.Lines)
            {
                string key = excelLine.FromAccountName;

                // Check if we have any lines for this account
                if (!qbLinesByAccount.ContainsKey(key) || qbLinesByAccount[key].Count == 0)
                    return false;

                // Try to find a matching line with the same amount (approximately)
                bool foundMatch = false;

                foreach (var qbLine in qbLinesByAccount[key].ToList())
                {
                    if (Math.Abs(qbLine.Amount - excelLine.Amount) <= tolerance)
                    {
                        // Found a match, remove it from the list to avoid double-matching
                        qbLinesByAccount[key].Remove(qbLine);
                        foundMatch = true;
                        break;
                    }
                }

                if (!foundMatch)
                    return false;
            }

            // If we've made it here, all checks passed
            return true;
        }

        /// <summary>
        /// Prints the differences between two MiscIncome records
        /// </summary>
        /// <param name="qbIncome">MiscIncome record from QuickBooks</param>
        /// <param name="excelIncome">MiscIncome record from Excel</param>
        private static void PrintMiscIncomeDifferences(MiscIncome qbIncome, MiscIncome excelIncome)
        {
            Console.WriteLine($"  Differences for MiscIncome record with Memo '{excelIncome.Memo}':");

            // Basic properties
            if (qbIncome.DepositToAccount != excelIncome.DepositToAccount)
                Console.WriteLine($"  - Deposit Account: QB='{qbIncome.DepositToAccount}', Excel='{excelIncome.DepositToAccount}'");

            // Total amount
            if (Math.Abs(qbIncome.TotalAmount - excelIncome.TotalAmount) > 0.01)
                Console.WriteLine($"  - Total Amount: QB={qbIncome.TotalAmount:C}, Excel={excelIncome.TotalAmount:C}");

            // Line counts
            if (qbIncome.Lines.Count != excelIncome.Lines.Count)
                Console.WriteLine($"  - Line Count: QB={qbIncome.Lines.Count}, Excel={excelIncome.Lines.Count}");

            // Create summary of lines by account for easier comparison
            var qbLinesByAccount = qbIncome.Lines
                .GroupBy(l => l.FromAccountName)
                .ToDictionary(g => g.Key, g => g.Sum(l => l.Amount));

            var excelLinesByAccount = excelIncome.Lines
                .GroupBy(l => l.FromAccountName)
                .ToDictionary(g => g.Key, g => g.Sum(l => l.Amount));

            // Get all account names from both records
            var allAccounts = new HashSet<string>(qbLinesByAccount.Keys.Concat(excelLinesByAccount.Keys));

            // Compare lines by account
            foreach (var account in allAccounts)
            {
                double qbAmount = qbLinesByAccount.ContainsKey(account) ? qbLinesByAccount[account] : 0;
                double excelAmount = excelLinesByAccount.ContainsKey(account) ? excelLinesByAccount[account] : 0;

                if (Math.Abs(qbAmount - excelAmount) > 0.01)
                {
                    Console.WriteLine($"  - Account '{account}': QB={qbAmount:C}, Excel={excelAmount:C}");
                }
            }
        }

        /// <summary>
        /// Gets only the NEW records that need to be added to QuickBooks
        /// </summary>
        /// <param name="excelIncomes">List of MiscIncome records from Excel</param>
        /// <returns>List of NEW MiscIncome records not found in QuickBooks</returns>
        public static List<MiscIncome> GetNewRecords(List<MiscIncome> excelIncomes)
        {
            List<MiscIncome> qbIncomes = MiscIncomeReader.QueryAllMiscIncomes();
            var qbIncomesDict = qbIncomes.ToDictionary(i => i.Memo, i => i);

            // Filter for only new records
            return excelIncomes.Where(income => !qbIncomesDict.ContainsKey(income.Memo)).ToList();
        }
    }
}