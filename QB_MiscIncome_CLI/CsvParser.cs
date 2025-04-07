using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using QB_MiscIncome_Lib;

namespace QB_MiscIncome_CLI
{
    /// <summary>
    /// Handles parsing of CSV deposit data file into MiscIncome objects
    /// </summary>
    public class CsvParser
    {
        /// <summary>
        /// Parses a CSV file and returns a list of MiscIncome objects
        /// </summary>
        /// <param name="filePath">Path to the CSV file</param>
        /// <returns>List of MiscIncome objects</returns>
        public static List<MiscIncome> ParseCsvFile(string filePath)
        {
            List<MiscIncome> records = new List<MiscIncome>();

            try
            {
                // Read all lines from the CSV file
                string[] lines = File.ReadAllLines(filePath);

                if (lines.Length <= 1)
                {
                    Console.WriteLine("CSV file is empty or contains only headers.");
                    return records;
                }

                // Skip header row and process each data row
                for (int i = 1; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var record = ParseCsvLine(line);
                    if (record != null)
                    {
                        records.Add(record);
                        Console.WriteLine($"Parsed record: {record.TxnID}");
                    }
                }

                Console.WriteLine($"Successfully parsed {records.Count} records from CSV file.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing CSV file: {ex.Message}");
            }

            return records;
        }

        /// <summary>
        /// Parses a single line from the CSV file into a MiscIncome object
        /// </summary>
        /// <param name="line">CSV line</param>
        /// <returns>MiscIncome object</returns>
        private static MiscIncome ParseCsvLine(string line)
        {
            try
            {
                // Split the CSV line using commas, but handle quoted values properly
                string[] fields = SplitCsvLine(line);

                // Check if we have at least the minimum number of fields (based on CSV header)
                if (fields.Length < 14)
                {
                    Console.WriteLine($"Warning: Line does not contain enough fields: {line}");
                    return null;
                }

                // Extract data from CSV fields
                // Assume CSV fields are in the following order:
                // 0=Parent ID, 1=Child ID, 4=Bank Date, 5=Customer, 6=Check Amount, 
                // 7=Tier 2 - Chart of Account ID, 8=Tier 2 - Chart of Account, 11=Tier 1 - Type

                string txnID = fields[0].Trim(); // Transaction ID
                DateTime depositDate = DateTime.Parse(fields[4].Trim()); // Deposit Date
                string memo = fields[2].Trim(); // Memo field (can be used for CompanyID)
                string depositToAccount = fields[3].Trim(); // Account to which the deposit was made
                double totalAmount = double.Parse(fields[6].Trim().Replace("$", "").Replace(",", "").Trim('"')); // Total Amount

                // Create a MiscIncome object
                MiscIncome income = new MiscIncome
                {
                    TxnID = txnID,
                    DepositDate = depositDate,
                    Memo = memo,
                    DepositToAccount = depositToAccount,
                    TotalAmount = totalAmount,
                    Lines = new List<MiscIncomeLine>()
                };

                // Add a MiscIncomeLine for each deposit line (e.g., customer info, account, etc.)
                MiscIncomeLine lineItem = new MiscIncomeLine
                {
                    ReceivedFromName = fields[5].Trim(), // Customer Name
                    FromAccountName = fields[8].Trim(), // From Account Name
                    Amount = double.Parse(fields[6].Trim().Replace("$", "").Replace(",", "").Trim('"')), // Amount for the line
                    Memo = fields[9].Trim() // Optional memo for this line (if available)
                };

                // Add the line item to the MiscIncome's Lines collection
                income.Lines.Add(lineItem);

                return income;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing CSV line: {ex.Message}");
                Console.WriteLine($"Line: {line}");
                return null;
            }
        }

        /// <summary>
        /// Splits a CSV line, handling quoted values correctly
        /// </summary>
        private static string[] SplitCsvLine(string line)
        {
            List<string> result = new List<string>();
            bool inQuotes = false;
            StringBuilder field = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    // Toggle in-quotes state
                    inQuotes = !inQuotes;
                    field.Append(c); // Keep quotes in the field value
                }
                else if (c == ',' && !inQuotes)
                {
                    // End of field
                    result.Add(field.ToString());
                    field.Clear();
                }
                else
                {
                    field.Append(c);
                }
            }

            // Add the last field
            result.Add(field.ToString());

            return result.ToArray();
        }
    }
}
