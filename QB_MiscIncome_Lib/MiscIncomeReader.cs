using System;
using System.Collections.Generic;
using System.Linq;
using QBFC16Lib;
using Serilog;

namespace QB_MiscIncome_Lib
{
    /// <summary>
    /// Class for querying and retrieving deposit transactions from QuickBooks
    /// </summary>
    public class MiscIncomeReader
    {
        /// <summary>
        /// Queries all deposit transactions from QuickBooks
        /// </summary>
        /// <returns>A list of MiscIncome objects</returns>
        public static List<MiscIncome> QueryAllMiscIncomes()
        {
            List<MiscIncome> results = new List<MiscIncome>();
            QBSessionManager qbSession = null;
            bool sessionBegun = false;
            bool connectionOpen = false;

            try
            {
                // Initialize the session manager
                qbSession = new QBSessionManager();

                // Open connection to QuickBooks using the app name from AppConfig
                qbSession.OpenConnection("", AppConfig.QB_APP_NAME);
                connectionOpen = true;

                qbSession.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                // Create request message set
                IMsgSetRequest requestMsgSet = qbSession.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                // Create deposit query
                IDepositQuery depositQuery = requestMsgSet.AppendDepositQueryRq();
                depositQuery.IncludeLineItems.SetValue(true);

                // Send request to QuickBooks
                IMsgSetResponse responseMsgSet = qbSession.DoRequests(requestMsgSet);

                // Process the response
                results = ProcessDepositQueryResponse(responseMsgSet);

                Log.Information($"Retrieved {results.Count} deposit transactions from QuickBooks");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error querying deposit transactions from QuickBooks");
                throw;
            }
            finally
            {
                // Clean up the session
                if (sessionBegun)
                {
                    qbSession.EndSession();
                }

                if (connectionOpen)
                {
                    qbSession.CloseConnection();
                }
            }

            return results;
        }

        /// <summary>
        /// Process the deposit query response from QuickBooks
        /// </summary>
        private static List<MiscIncome> ProcessDepositQueryResponse(IMsgSetResponse responseMsgSet)
        {
            List<MiscIncome> deposits = new List<MiscIncome>();

            if (responseMsgSet == null) return deposits;

            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return deposits;

            // Check for valid response
            if (responseList.Count > 0)
            {
                IResponse response = responseList.GetAt(0);

                // Check status code (0 = OK, >0 is warning)
                if (response.StatusCode >= 0)
                {
                    if (response.Detail != null)
                    {
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if (responseType == ENResponseType.rtDepositQueryRs)
                        {
                            // Get the deposit return list
                            IDepositRetList depositRetList = (IDepositRetList)response.Detail;

                            if (depositRetList != null)
                            {
                                // Process each deposit transaction
                                for (int i = 0; i < depositRetList.Count; i++)
                                {
                                    IDepositRet depositRet = depositRetList.GetAt(i);
                                    MiscIncome deposit = ConvertToMiscIncome(depositRet);
                                    deposits.Add(deposit);
                                }
                            }
                        }
                    }
                }
                else
                {
                    Log.Error($"QuickBooks error: {response.StatusMessage} (Code {response.StatusCode})");
                }
            }

            return deposits;
        }

        /// <summary>
        /// Convert a QuickBooks deposit transaction to our MiscIncome model
        /// </summary>
        private static MiscIncome ConvertToMiscIncome(IDepositRet depositRet)
        {
            MiscIncome deposit = new MiscIncome
            {
                TxnID = depositRet.TxnID.GetValue(),
                DepositDate = depositRet.TxnDate.GetValue(),
                DepositToAccount = depositRet.DepositToAccountRef.FullName?.GetValue() ?? "",
                TotalAmount = depositRet.DepositTotal?.GetValue() ?? 0.0
            };

            // Get memo field (which stores CompanyID)
            if (depositRet.Memo != null)
            {
                deposit.Memo = depositRet.Memo.GetValue();
            }

            // Process each line in the deposit
            if (depositRet.DepositLineRetList != null)
            {
                for (int j = 0; j < depositRet.DepositLineRetList.Count; j++)
                {
                    IDepositLineRet lineRet = depositRet.DepositLineRetList.GetAt(j);
                    MiscIncomeLine line = new MiscIncomeLine();

                    // Get line amount
                    if (lineRet.Amount != null)
                    {
                        line.Amount = lineRet.Amount.GetValue();
                    }

                    // Get account info
                    if (lineRet.AccountRef != null)
                    {
                        if (lineRet.AccountRef.ListID != null)
                        {
                            line.FromAccountListID = lineRet.AccountRef.ListID.GetValue();
                        }

                        if (lineRet.AccountRef.FullName != null)
                        {
                            line.FromAccountName = lineRet.AccountRef.FullName.GetValue();
                        }
                    }

                    // Get customer/entity info (ReceivedFrom)
                    if (lineRet.EntityRef != null)
                    {
                        if (lineRet.EntityRef.ListID != null)
                        {
                            line.ReceivedFromListID = lineRet.EntityRef.ListID.GetValue();
                        }

                        if (lineRet.EntityRef.FullName != null)
                        {
                            line.ReceivedFromName = lineRet.EntityRef.FullName.GetValue();
                        }
                    }

                    // Get line memo
                    if (lineRet.Memo != null)
                    {
                        line.Memo = lineRet.Memo.GetValue();
                    }

                    deposit.Lines.Add(line);
                }
            }

            return deposit;
        }
    }
}