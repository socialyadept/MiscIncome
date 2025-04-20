using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Xunit;
using QBFC16Lib;
using QB_MiscIncome_Lib; // your library that defines MiscIncome, etc.

namespace QB_MiscIncome_Test
{
    [Collection("Sequential Tests")]
    public class MiscIncomeAdderTests
    {
        [Fact]
        public void AddMiscIncomes_ShouldAssignTxnIDs_AndAppearInQB()
        {
            // Track newly added deposits, so we can delete them afterward
            var createdDepositTxnIDs = new List<string>();

            // Track newly added customers, so we can delete them afterward
            var createdCustomerListIDs = new List<string>();

            // We'll build up a list of MiscIncome objects,
            // each deposit references a known customer in QuickBooks.
            var incomesToAdd = new List<MiscIncome>();

            // Create random test data
            try
            {
                // 1) Create any needed customers in QB
                using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    //qbSession.BeginSession(); // only if your QuickBooksSession requires it
                    for (int i = 0; i < 2; i++)
                    {
                        string custName = $"TestCust_{Guid.NewGuid().ToString("N").Substring(0,6)}";
                        string custListID = AddCustomer(qbSession, custName);
                        createdCustomerListIDs.Add(custListID);

                        // 2) Build a deposit referencing the newly-created customer
                        var deposit = new MiscIncome
                        {
                            DepositDate = DateTime.Today.AddDays(-i),
                            Memo = (100 + i).ToString(), // storing a numeric "CompanyID" in the memo
                            DepositToAccount = "Checking",
                            TotalAmount = 50.0 + (25 * i),
                            Lines = new List<MiscIncomeLine>
                            {
                                new MiscIncomeLine
                                {
                                    // We reference the real customer we just created
                                    ReceivedFromListID = custListID,
                                    ReceivedFromName = custName,
                                    FromAccountName = "Sales",
                                    Amount = 30.0 + (10 * i),
                                    Memo = "Line memo #1"
                                },
                                new MiscIncomeLine
                                {
                                    ReceivedFromListID = custListID,
                                    ReceivedFromName = custName,
                                    FromAccountName = "Shipping and Delivery Income",
                                    Amount = 20.0,
                                    Memo = "Line memo #2"
                                }
                            }
                        };
                        incomesToAdd.Add(deposit);
                    }
                }

                // 3) Call your adder code to add these deposits to QB
                using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    var adder = new MiscIncomeAdder(qbSession);
                    adder.AddMiscIncomes(incomesToAdd);
                }

                // 4) Verify that each deposit has a valid TxnID
                foreach (var inc in incomesToAdd)
                {
                    Assert.False(string.IsNullOrWhiteSpace(inc.TxnID),
                        $"Expected TxnID to be set after adding deposit. Memo={inc.Memo}");
                    createdDepositTxnIDs.Add(inc.TxnID);
                }

                // 5) For each deposit, query QB by TxnID to ensure it truly exists
                foreach (var inc in incomesToAdd)
                {
                    var qbDeposit = QueryDepositByTxnID(inc.TxnID);
                    Assert.NotNull(qbDeposit);  // If null, deposit not found in QB
                    Assert.Equal(inc.TxnID, qbDeposit!.TxnID);

                    // Optional deeper checks
                    Assert.Equal(inc.Memo, qbDeposit.Memo);
                    Assert.Equal(inc.DepositDate.Date, qbDeposit.DepositDate.Date);
                    Assert.Equal(inc.DepositToAccount, qbDeposit.DepositToAccount);
                    Assert.Equal(inc.Lines.Count, qbDeposit.Lines.Count);
                }
            }
            finally
            {
                // 6) Cleanup: remove deposits first
                using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    foreach (var txnID in createdDepositTxnIDs)
                    {
                        DeleteDeposit(qbSession, txnID);
                    }
                }

                // 7) Cleanup: then remove customers
                using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    foreach (var custID in createdCustomerListIDs)
                    {
                        DeleteListObject(qbSession, custID, ENListDelType.ldtCustomer);
                    }
                }
            }
        }

        // ---------------------------------------------------------------------
        // HELPER: CREATE CUSTOMER
        // ---------------------------------------------------------------------
        private string AddCustomer(QuickBooksSession qbSession, string customerName)
        {
            IMsgSetRequest request = qbSession.CreateRequestSet();
            var custAdd = request.AppendCustomerAddRq();
            custAdd.Name.SetValue(customerName);

            IMsgSetResponse resp = qbSession.SendRequest(request);
            return ExtractCustomerListID(resp);
        }

        private string ExtractCustomerListID(IMsgSetResponse resp)
        {
            var list = resp.ResponseList;
            if (list == null || list.Count == 0)
                throw new Exception("No response from CustomerAdd.");

            IResponse r = list.GetAt(0);
            if (r.StatusCode != 0)
                throw new Exception($"CustomerAdd failed: {r.StatusMessage}");

            var custRet = r.Detail as ICustomerRet;
            if (custRet == null)
                throw new Exception("No ICustomerRet returned.");

            return custRet.ListID.GetValue();
        }

        // ---------------------------------------------------------------------
        // HELPER: QUERY DEPOSIT
        // ---------------------------------------------------------------------
        private MiscIncome? QueryDepositByTxnID(string txnID)
        {
            using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
            {
                IMsgSetRequest request = qbSession.CreateRequestSet();
                request.Attributes.OnError = ENRqOnError.roeContinue;

                // Build a DepositQuery request, searching by TxnID
                var depositQuery = request.AppendDepositQueryRq();
                depositQuery.IncludeLineItems.SetValue(true);
                depositQuery.ORDepositQuery.TxnIDList.Add(txnID);

                IMsgSetResponse response = qbSession.SendRequest(request);
                IResponseList responseList = response.ResponseList;
                if (responseList == null || responseList.Count == 0)
                    return null;

                IResponse resp = responseList.GetAt(0);
                if (resp.StatusCode != 0)
                {
                    Debug.WriteLine($"DepositQuery failed: {resp.StatusMessage}");
                    return null;
                }

                // If we got a deposit in the result, convert to a MiscIncome model
                var depositRetList = resp.Detail as IDepositRetList; 
                if (depositRetList == null || depositRetList.Count == 0)
                    return null;

                var depositRet = depositRetList.GetAt(0);
                if (depositRet == null)
                    return null;

                return ConvertDepositRetToMiscIncome(depositRet);
            }
        }

        private MiscIncome ConvertDepositRetToMiscIncome(IDepositRet depositRet)
        {
            var inc = new MiscIncome
            {
                TxnID = depositRet.TxnID.GetValue(),
                DepositDate = depositRet.TxnDate.GetValue(),
                DepositToAccount = depositRet.DepositToAccountRef?.FullName?.GetValue() ?? "",
                Memo = depositRet.Memo?.GetValue() ?? "",
                TotalAmount = depositRet.DepositTotal?.GetValue() ?? 0.0
            };

            if (depositRet.DepositLineRetList != null)
            {
                for (int i = 0; i < depositRet.DepositLineRetList.Count; i++)
                {
                    var lineRet = depositRet.DepositLineRetList.GetAt(i);
                    if (lineRet == null) continue;

                    var line = new MiscIncomeLine
                    {
                        ReceivedFromListID = lineRet.EntityRef?.ListID?.GetValue() ?? "",
                        ReceivedFromName = lineRet.EntityRef?.FullName?.GetValue() ?? "",
                        FromAccountListID = lineRet.AccountRef?.ListID?.GetValue() ?? "",
                        FromAccountName = lineRet.AccountRef?.FullName?.GetValue() ?? "",
                        Memo = lineRet.Memo?.GetValue() ?? "",
                        Amount = lineRet.Amount?.GetValue() ?? 0.0
                    };

                    inc.Lines.Add(line);
                }
            }
            return inc;
        }

        // ---------------------------------------------------------------------
        // HELPER: DELETE DEPOSIT
        // ---------------------------------------------------------------------
        private void DeleteDeposit(QuickBooksSession qbSession, string txnID)
        {
            IMsgSetRequest request = qbSession.CreateRequestSet();
            var delReq = request.AppendTxnDelRq();
            delReq.TxnDelType.SetValue(ENTxnDelType.tdtDeposit);
            delReq.TxnID.SetValue(txnID);

            IMsgSetResponse resp = qbSession.SendRequest(request);
            CheckForError(resp, $"Deleting Deposit {txnID}");
        }

        // ---------------------------------------------------------------------
        // HELPER: DELETE LIST OBJECT (e.g. CUSTOMER)
        // ---------------------------------------------------------------------
        private void DeleteListObject(QuickBooksSession qbSession, string listID, ENListDelType listDelType)
        {
            IMsgSetRequest request = qbSession.CreateRequestSet();
            IListDel listDel = request.AppendListDelRq();
            listDel.ListDelType.SetValue(listDelType);
            listDel.ListID.SetValue(listID);

            IMsgSetResponse resp = qbSession.SendRequest(request);
            CheckForError(resp, $"Deleting {listDelType} {listID}");
        }

        // ---------------------------------------------------------------------
        // HELPER: CHECK FOR ERRORS
        // ---------------------------------------------------------------------
        private void CheckForError(IMsgSetResponse resp, string context)
        {
            if (resp?.ResponseList == null || resp.ResponseList.Count == 0)
                return;

            var r = resp.ResponseList.GetAt(0);
            if (r.StatusCode != 0)
            {
                throw new Exception($"Error {context}: {r.StatusMessage} (StatusCode={r.StatusCode})");
            }
            else
            {
                Debug.WriteLine($"OK: {context}");
            }
        }
    }
}
