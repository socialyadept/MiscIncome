using System.Diagnostics;
using Serilog;
using QBFC16Lib;
using static QB_MiscIncome_Test.CommonMethods; // Reuse or adapt your shared helpers

namespace QB_MiscIncome_Test
{
    [Collection("Sequential Tests")]
    public class MiscIncomeReaderTests
    {
        private const int CUSTOMER_COUNT = 2;

        [Fact]
        public void CreateAndDelete_Customers_Deposits()
        {
            var createdCustomerListIDs = new List<string>();
            var createdDepositTxnIDs = new List<string>();

            // We'll store test data so we can verify it after querying
            var randomCustomerNames = new List<string>();
            var depositTestData = new List<DepositTestInfo>();

            try
            {
                // 1) Clean logs
                EnsureLogFileClosed();
                DeleteOldLogFiles();
                ResetLogger();

                // 2) Create 2 customers
                using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    for (int i = 0; i < CUSTOMER_COUNT; i++)
                    {
                        string custName = "RandCust_" + Guid.NewGuid().ToString("N").Substring(0, 6);
                        string custListID = AddCustomer(qbSession, custName);

                        createdCustomerListIDs.Add(custListID);
                        randomCustomerNames.Add(custName);
                    }
                }

                // 3) Create 1 deposit per customer. Each deposit has 2 lines:
                //    - line 1: FromAccount = "Sales"
                //    - line 2: FromAccount = "Shipping and Delivery Income"
                //    The deposit goes into "Checking" or any account you want.
                //    We store a numeric CompanyID in the Memo, e.g. 100, 101, ...
                using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    for (int i = 0; i < CUSTOMER_COUNT; i++)
                    {
                        string custListID = createdCustomerListIDs[i];
                        string custName = randomCustomerNames[i];

                        int companyID = 100 + i;  // numeric
                        double line1Amount = 25.0 + i; // vary a bit
                        double line2Amount = 10.0 + i; // vary a bit
                        DateTime depositDate = DateTime.Today;

                        string depositTxnID = AddDepositWithTwoLines(
                            qbSession,
                            custListID,
                            custName,
                            companyID,
                            depositDate,
                            line1Amount,
                            line2Amount
                        );

                        createdDepositTxnIDs.Add(depositTxnID);

                        // Save for final asserts
                        depositTestData.Add(new DepositTestInfo
                        {
                            TxnID = depositTxnID,
                            CustomerName = custName,
                            CompanyID = companyID,
                            DepositDate = depositDate,
                            LineAccounts = new List<string> { "Sales", "Shipping and Delivery Income" },
                            LineAmounts = new List<double> { line1Amount, line2Amount }
                        });
                    }
                }

                // 4) Query & verify
                var allDeposits = MiscIncomeReader.QueryAllMiscIncomes();
                // Similar to your InvoiceReader or PaymentReader, but returning a "Deposit" model.

                foreach (var d in depositTestData)
                {
                    var matchingDep = allDeposits.FirstOrDefault(x => x.TxnID == d.TxnID);
                    Assert.NotNull(matchingDep);

                    // Check numeric CompanyID in Memo
                    // e.g. if your deposit has a property "Memo" that's just the numeric string
                    // If so, we might do: 
                    Assert.Equal(d.CompanyID.ToString(), matchingDep.Memo);

                    // Check deposit date
                    Assert.Equal(d.DepositDate.Date, matchingDep.DepositDate.Date);

                    // Check each line
                    // e.g. if your deposit model has a list of lines with (ReceivedFrom, FromAccount, Amount)
                    Assert.Equal(2, matchingDep.Lines.Count);

                    // The order of lines might vary, so you might do more flexible checks. 
                    // We'll assume they appear in the same order we created them:
                    Assert.Equal(d.LineAccounts[0], matchingDep.Lines[0].FromAccountName);
                    Assert.Equal(d.LineAccounts[1], matchingDep.Lines[1].FromAccountName);

                    Assert.Equal(d.LineAmounts[0], matchingDep.Lines[0].Amount);
                    Assert.Equal(d.LineAmounts[1], matchingDep.Lines[1].Amount);

                    // If you also store "ReceivedFrom" as the customer:
                    Assert.Equal(d.CustomerName, matchingDep.Lines[0].ReceivedFromName);
                    Assert.Equal(d.CustomerName, matchingDep.Lines[1].ReceivedFromName);
                }
            }
            finally
            {
                // 5) Cleanup: remove deposits first, then customers
                using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    foreach (var depID in createdDepositTxnIDs)
                    {
                        DeleteDeposit(qbSession, depID);
                    }
                }

                using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    foreach (var custID in createdCustomerListIDs)
                    {
                        DeleteListObject(qbSession, custID, ENListDelType.ldtCustomer);
                    }
                }
            }
        }

        //-------------------------------------------------------------------------
        // "MAKE DEPOSIT" CREATION
        //-------------------------------------------------------------------------

        private string AddDepositWithTwoLines(
            QuickBooksSession qbSession,
            string customerListID,
            string customerName,
            int companyID,
            DateTime depositDate,
            double line1Amount,
            double line2Amount
        )
        {
            IMsgSetRequest request = qbSession.CreateRequestSet();
            // The top-level deposit add
            var depositAdd = request.AppendDepositAddRq();

            // The deposit goes "TO" some account (like Checking)
            // Adjust to match your QuickBooks "Deposit to" account
            depositAdd.DepositToAccountRef.FullName.SetValue("Checking");

            // The numeric CompanyID in Memo
            depositAdd.Memo.SetValue(companyID.ToString());
            depositAdd.TxnDate.SetValue(depositDate);

            // 1) First line
            //    ReceivedFrom = this customer
            //    FromAccountRef = "Sales"
            //    Amount = line1Amount
            //    If you also want to store the name in the line's memo, you can do that.
            var line1 = depositAdd.DepositLineAddList.Append();
            line1.ReceivedFromRef.ListID.SetValue(customerListID);
            line1.AccountRef.FullName.SetValue("Sales");
            line1.Amount.SetValue(line1Amount);

            // 2) Second line
            //    ReceivedFrom = same customer
            //    FromAccountRef = "Shipping and Delivery Income"
            //    Amount = line2Amount
            var line2 = depositAdd.DepositLineAddList.Append();
            line2.ReceivedFromRef.ListID.SetValue(customerListID);
            line2.AccountRef.FullName.SetValue("Shipping and Delivery Income");
            line2.Amount.SetValue(line2Amount);

            // Send to QB
            var resp = qbSession.SendRequest(request);
            return ExtractDepositTxnID(resp);
        }

        //-------------------------------------------------------------------------
        // CREATE CUSTOMER
        //-------------------------------------------------------------------------

        private string AddCustomer(QuickBooksSession qbSession, string customerName)
        {
            IMsgSetRequest request = qbSession.CreateRequestSet();
            var custAdd = request.AppendCustomerAddRq();
            custAdd.Name.SetValue(customerName);

            var resp = qbSession.SendRequest(request);
            return ExtractCustomerListID(resp);
        }

        //-------------------------------------------------------------------------
        // DELETING
        //-------------------------------------------------------------------------

        private void DeleteDeposit(QuickBooksSession qbSession, string txnID)
        {
            // For removing a deposit, we use TxnDelRq with tdtDeposit
            IMsgSetRequest request = qbSession.CreateRequestSet();
            var delReq = request.AppendTxnDelRq();
            delReq.TxnDelType.SetValue(ENTxnDelType.tdtDeposit);
            delReq.TxnID.SetValue(txnID);

            var resp = qbSession.SendRequest(request);
            CheckForError(resp, $"Deleting Deposit {txnID}");
        }

        private void DeleteListObject(QuickBooksSession qbSession, string listID, ENListDelType listDelType)
        {
            IMsgSetRequest request = qbSession.CreateRequestSet();
            IListDel listDel = request.AppendListDelRq();
            listDel.ListDelType.SetValue(listDelType);
            listDel.ListID.SetValue(listID);

            var resp = qbSession.SendRequest(request);
            CheckForError(resp, $"Deleting {listDelType} {listID}");
        }

        //-------------------------------------------------------------------------
        // EXTRACTORS
        //-------------------------------------------------------------------------

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

        private string ExtractDepositTxnID(IMsgSetResponse resp)
        {
            var list = resp.ResponseList;
            if (list == null || list.Count == 0)
                throw new Exception("No response from DepositAddRq.");

            IResponse r = list.GetAt(0);
            if (r.StatusCode != 0)
                throw new Exception($"DepositAdd failed: {r.StatusMessage}");

            // The deposit ret object is IDepositRet
            var depositRet = r.Detail as IDepositRet;
            if (depositRet == null)
                throw new Exception("No IDepositRet returned.");

            return depositRet.TxnID.GetValue();
        }

        //-------------------------------------------------------------------------
        // ERROR HANDLER
        //-------------------------------------------------------------------------

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

        //-------------------------------------------------------------------------
        // TEST DATA MODEL
        //-------------------------------------------------------------------------

        private class DepositTestInfo
        {
            public string TxnID { get; set; } = "";
            public string CustomerName { get; set; } = "";
            public int CompanyID { get; set; } // numeric
            public DateTime DepositDate { get; set; }
            public List<string> LineAccounts { get; set; } = new();
            public List<double> LineAmounts { get; set; } = new();
        }
    }
}
