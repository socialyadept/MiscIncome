using System;
using System.Collections.Generic;
using System.Linq;
using QBFC16Lib;
using Serilog;

namespace QB_MiscIncome_Lib
{
    /// <summary>
    /// Class responsible for adding MiscIncome objects (deposits) to QuickBooks
    /// </summary>
    public class MiscIncomeAdder
    {
        private readonly QuickBooksSession _qbSession;

        /// <summary>
        /// Initializes a new instance of the MiscIncomeAdder class
        /// </summary>
        /// <param name="qbSession">The QuickBooks session to use</param>
        public MiscIncomeAdder(QuickBooksSession qbSession)
        {
            _qbSession = qbSession ?? throw new ArgumentNullException(nameof(qbSession));
        }

        /// <summary>
        /// Adds multiple MiscIncome objects to QuickBooks
        /// </summary>
        /// <param name="miscIncomes">The list of MiscIncome objects to add</param>
        public void AddMiscIncomes(List<MiscIncome> miscIncomes)
        {
            if (miscIncomes == null || !miscIncomes.Any())
            {
                Log.Warning("No MiscIncome objects to add.");
                return;
            }

            Log.Information($"Adding {miscIncomes.Count} MiscIncome objects to QuickBooks");

            // Process each MiscIncome object
            foreach (var income in miscIncomes)
            {
                try
                {
                    AddMiscIncome(income);
                    Log.Information($"Successfully added MiscIncome with memo: {income.Memo}");
                }
                catch (Exception ex)
                {
                    Log.Error(ex, $"Error adding MiscIncome with memo: {income.Memo}");
                    throw;
                }
            }
        }

        /// <summary>
        /// Adds a single MiscIncome object to QuickBooks
        /// </summary>
        /// <param name="income">The MiscIncome object to add</param>
        private void AddMiscIncome(MiscIncome income)
        {
            // Create request message set
            IMsgSetRequest requestMsgSet = _qbSession.CreateRequestSet();

            // Create deposit add request
            IDepositAdd depositAdd = requestMsgSet.AppendDepositAddRq();

            // Set deposit date
            depositAdd.TxnDate.SetValue(income.DepositDate);

            // Set memo (CompanyID)
            if (!string.IsNullOrWhiteSpace(income.Memo))
            {
                depositAdd.Memo.SetValue(income.Memo);
            }

            // Set deposit account
            if (!string.IsNullOrWhiteSpace(income.DepositToAccount))
            {
                depositAdd.DepositToAccountRef.FullName.SetValue(income.DepositToAccount);
            }

            // Add deposit lines
            if (income.Lines != null && income.Lines.Any())
            {
                foreach (var line in income.Lines)
                {
                    // Add a new deposit line
                    IDepositLineAdd lineAdd = depositAdd.DepositLineAddList.Append();

                    // Using the ORDepositLineAdd.DepositInfo structure
                    // This is the correct way to set line properties based on the QuickBooks SDK

                    // Set the entity (customer/vendor)
                    if (!string.IsNullOrWhiteSpace(line.ReceivedFromListID))
                    {
                        lineAdd.ORDepositLineAdd.DepositInfo.EntityRef.ListID.SetValue(line.ReceivedFromListID);
                    }
                    else if (!string.IsNullOrWhiteSpace(line.ReceivedFromName))
                    {
                        lineAdd.ORDepositLineAdd.DepositInfo.EntityRef.FullName.SetValue(line.ReceivedFromName);
                    }

                    // Set the account
                    if (!string.IsNullOrWhiteSpace(line.FromAccountName))
                    {
                        lineAdd.ORDepositLineAdd.DepositInfo.AccountRef.FullName.SetValue(line.FromAccountName);
                    }

                    // Set the memo
                    if (!string.IsNullOrWhiteSpace(line.Memo))
                    {
                        lineAdd.ORDepositLineAdd.DepositInfo.Memo.SetValue(line.Memo);
                    }

                    // Set the amount
                    lineAdd.ORDepositLineAdd.DepositInfo.Amount.SetValue(line.Amount);
                }
            }

            // Send request to QuickBooks
            IMsgSetResponse responseMsgSet = _qbSession.SendRequest(requestMsgSet);

            // Process the response
            if (responseMsgSet == null)
            {
                throw new Exception("No response received from QuickBooks");
            }

            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null || responseList.Count == 0)
            {
                throw new Exception("Empty response list from QuickBooks");
            }

            // Check the status of the response
            IResponse response = responseList.GetAt(0);
            if (response.StatusCode < 0)
            {
                throw new Exception($"QuickBooks error: {response.StatusMessage} (Code {response.StatusCode})");
            }

            // Extract the TxnID from the response and assign it to the MiscIncome object
            if (response.Detail != null && response.Type.GetValue() == (int)ENResponseType.rtDepositAddRs)
            {
                IDepositRet depositRet = (IDepositRet)response.Detail;
                income.TxnID = depositRet.TxnID.GetValue();
            }
            else
            {
                throw new Exception("Unexpected response type from QuickBooks");
            }
        }
    }
}