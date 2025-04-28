using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;
using QB_MiscIncome_Lib;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace QB_MiscIncome_CLI
{
    /// <summary>
    /// Handles parsing of Excel (.xlsx) MiscIncome data file
    /// </summary>
    public static class ExcelParser
    {
        /// <summary>
        /// Parses an Excel file and returns a list of MiscIncome objects
        /// </summary>
        /// <param name="filePath">Path to the Excel file</param>
        /// <returns>List of MiscIncome objects</returns>
        public static List<MiscIncome> ParseExcelFile(string filePath)
        {
            List<MiscIncome> records = new List<MiscIncome>();

            try
            {
                Console.WriteLine($"Reading Excel file: {filePath}");

                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    Worksheet worksheet = worksheetPart.Worksheet;

                    // Get the shared string table
                    SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
                    SharedStringTable sharedStringTable = stringTablePart?.SharedStringTable;

                    // Get the sheet data
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData == null || !sheetData.Any())
                    {
                        Console.WriteLine("Excel file is empty or contains no data.");
                        return records;
                    }

                    // Get all rows
                    var rows = sheetData.Elements<Row>().ToList();
                    if (rows.Count <= 1)
                    {
                        Console.WriteLine("Excel file contains only headers or is empty.");
                        return records;
                    }

                    // First row should be headers
                    var headerRow = rows[0];

                    // Create a mapping of column indices to header names
                    Dictionary<string, int> columnMapping = new Dictionary<string, int>();

                    // Process header cells
                    foreach (Cell cell in headerRow.Elements<Cell>())
                    {
                        if (cell.CellReference == null)
                            continue;

                        string cellReference = cell.CellReference.Value;
                        string columnName = GetColumnName(cellReference);
                        string cellValue = GetCellValue(cell, sharedStringTable);

                        Console.WriteLine($"Header column {columnName} = '{cellValue}'");

                        // Store column letter to header name mapping
                        columnMapping[cellValue] = GetColumnIndex(columnName);
                    }

                    // Verify required columns exist
                    if (!columnMapping.ContainsKey("Child ID") ||
                        !columnMapping.ContainsKey("Check Amount") ||
                        !columnMapping.ContainsKey("Tier 2 - Chart of Account") ||
                        !columnMapping.ContainsKey("Tier 1 - Type"))
                    {
                        Console.WriteLine("Error: Required columns not found in Excel header.");
                        return records;
                    }

                    // Group data by Child ID to create separate MiscIncome records
                    Dictionary<string, MiscIncome> incomeByChildID = new Dictionary<string, MiscIncome>();

                    // Process data rows
                    for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++)
                    {
                        var row = rows[rowIndex];

                        Console.WriteLine($"Processing row {rowIndex}, RowIndex={row.RowIndex?.Value}");

                        try
                        {
                            // Create a dictionary to store the values for this row
                            Dictionary<string, string> rowValues = new Dictionary<string, string>();

                            // Process all cells in this row
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                if (cell.CellReference == null)
                                    continue;

                                string cellReference = cell.CellReference.Value;
                                string columnName = GetColumnName(cellReference);
                                string cellValue = GetCellValue(cell, sharedStringTable);

                                // For each header, check if this cell is in that column
                                foreach (var headerMapping in columnMapping)
                                {
                                    string headerText = headerMapping.Key;
                                    int columnIndex = headerMapping.Value;

                                    if (GetColumnIndex(columnName) == columnIndex)
                                    {
                                        rowValues[headerText] = cellValue;
                                        Console.WriteLine($"  Cell {cellReference} = '{cellValue}' (Header: {headerText})");
                                    }
                                }
                            }

                            // Check if we have all required values
                            if (!rowValues.ContainsKey("Child ID") || string.IsNullOrWhiteSpace(rowValues["Child ID"]))
                            {
                                Console.WriteLine("  Missing Child ID, skipping row");
                                continue;
                            }

                            if (!rowValues.ContainsKey("Check Amount") || string.IsNullOrWhiteSpace(rowValues["Check Amount"]))
                            {
                                Console.WriteLine("  Missing Check Amount, skipping row");
                                continue;
                            }

                            // Extract Child ID
                            string childId = rowValues["Child ID"];

                            // Parse amount
                            string amountStr = rowValues["Check Amount"].Replace("$", "").Replace(",", "").Trim();
                            Console.WriteLine($"  Amount string: '{amountStr}'");

                            if (!double.TryParse(amountStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double amount))
                            {
                                Console.WriteLine($"  Warning: Could not parse amount '{amountStr}'");
                                continue;
                            }

                            // Get account information based on Tier 1 and Tier 2
                            string accountName;
                            string accountType = rowValues.ContainsKey("Tier 1 - Type")
                                ? rowValues["Tier 1 - Type"]
                                : "Income"; // Default to Income

                            string chartOfAccount = rowValues.ContainsKey("Tier 2 - Chart of Account")
                                ? rowValues["Tier 2 - Chart of Account"]
                                : string.Empty;

                            // Determine account reference based on account type
                            switch (accountType)
                            {
                                case "Income":
                                    accountName = "Sales";
                                    break;
                                case "Expense":
                                    accountName = "Automobile Expense";
                                    break;
                                case "Equity":
                                    accountName = "Shareholder Distributions";
                                    break;
                                case "Other Income":
                                    // For Other Income, use the specific account name
                                    if (chartOfAccount == "Rental")
                                        accountName = "Rental";
                                    else
                                        accountName = "Misc Credits";
                                    break;
                                default:
                                    accountName = "Sales"; // Default to Sales
                                    break;
                            }

                            // Get line memo 
                            string lineMemo = string.Empty;
                            if (rowValues.ContainsKey("Description"))
                            {
                                lineMemo = rowValues["Description"];
                            }
                            else if (rowValues.ContainsKey("Memo"))
                            {
                                lineMemo = rowValues["Memo"];
                            }

                            // Get or create MiscIncome record for this Child ID
                            if (!incomeByChildID.TryGetValue(childId, out MiscIncome income))
                            {
                                // Create a new MiscIncome object
                                income = new MiscIncome
                                {
                                    DepositDate = DateTime.Now, // Use current date for deposit
                                    DepositToAccount = "Checking", // Default as specified
                                    Memo = childId // Use Child ID as Memo
                                };

                                incomeByChildID[childId] = income;
                            }

                            // Create a line item
                            var line = new MiscIncomeLine
                            {
                                ReceivedFromName = "Misc Income", // Default to "Misc Income" as specified
                                FromAccountName = accountName,
                                Amount = amount,
                                Memo = lineMemo
                            };

                            // Add line to income
                            income.Lines.Add(line);
                            income.TotalAmount += amount; // Update total amount

                            Console.WriteLine($"  Added line to MiscIncome {childId}: Amount={amount}, Account={accountName}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  Error processing row {rowIndex}: {ex.Message}");
                        }
                    }

                    // Convert the dictionary to list
                    records = incomeByChildID.Values.ToList();
                }

                Console.WriteLine($"\nSuccessfully parsed {records.Count} MiscIncome records with a total of " +
                                  $"{records.Sum(r => r.Lines.Count)} lines from Excel file.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing Excel file: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }

            return records;
        }

        /// <summary>
        /// Gets the value from a cell
        /// </summary>
        private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell == null)
                return string.Empty;

            string value = cell.InnerText;

            // If the cell has a data type attribute, we need to extract the value differently
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                // For shared strings, we need to look up the value in the shared string table
                if (sharedStringTable != null)
                {
                    int ssid = int.Parse(value);
                    if (ssid >= 0 && ssid < sharedStringTable.Count())
                    {
                        SharedStringItem ssi = (SharedStringItem)sharedStringTable.ElementAt(ssid);
                        if (ssi.Text != null)
                            value = ssi.Text.Text;
                        else if (ssi.InnerText != null)
                            value = ssi.InnerText;
                        else
                            value = string.Empty;
                    }
                }
            }
            else if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
            {
                value = value == "1" ? "TRUE" : "FALSE";
            }

            return value;
        }

        /// <summary>
        /// Extracts the column name from a cell reference (e.g., "A1" -> "A")
        /// </summary>
        private static string GetColumnName(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
                return string.Empty;

            // Return all characters until we hit a digit
            return new string(cellReference.TakeWhile(c => !char.IsDigit(c)).ToArray());
        }

        /// <summary>
        /// Converts column name to column index (e.g., "A" -> 0, "B" -> 1, "AA" -> 26)
        /// </summary>
        private static int GetColumnIndex(string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
                return -1;

            columnName = columnName.ToUpper();
            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum - 1; // Convert from 1-based to 0-based
        }
    }
}