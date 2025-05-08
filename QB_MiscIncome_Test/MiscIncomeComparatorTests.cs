// ─────────────────────────────────────────────────────────────────────────────
// QB_MiscIncome_Test | integration test exercising all 6 statuses
// ─────────────────────────────────────────────────────────────────────────────
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Serilog;
using QBFC16Lib;

using QB_Customers_Lib;          // CustomerAdder / CustomerReader
using QB_MiscIncome_Lib;         // MiscIncome, MiscIncomeComparator, MiscIncomeStatus
using static QB_Terms_Test.CommonMethods; // log-file helpers already in your project

namespace QB_MiscIncome_Test
{
    [Collection("Sequential Tests")]        // one QuickBooks session at a time
    public class MiscIncomeComparatorTests
    {
        private const int COMPANY_ID_START = 9100; // far from real IDs

        [Fact]
        public void CompareMiscIncomes_EndToEnd_AllStatusesCovered()
        {
            // ───────────────────────────────────────────────────────────
            // 1) Build all test data in memory
            // ───────────────────────────────────────────────────────────
            EnsureLogFileClosed();
            DeleteOldLogFiles();
            ResetLogger();

            var rnd             = new Random();
            var customers       = new List<Customer>();
            var miscIncomes     = new List<MiscIncome>();

            for (int i = 0; i < 5; i++)
            {
                // a) customer
                string custName = $"T_Cust_{Guid.NewGuid():N}".Substring(0, 10);
                customers.Add(new Customer(custName, $"Company_{i}"));

                // b) deposit linked to that customer
                var inc = new MiscIncome
                {
                    CompanyID        = COMPANY_ID_START + i,
                    DepositDate      = DateTime.Today,
                    DepositToAccount = "Checking",              // adjust if your file uses something else
                    TotalAmount      = Math.Round(rnd.NextDouble() * 100 + 50, 2),
                    Memo             = $"CID={COMPANY_ID_START + i}"
                };
                inc.Lines.Add(new MiscIncomeLine
                {
                    ReceivedFromName  = custName,
                    FromAccountName   = "Sales",
                    Amount            = inc.TotalAmount,
                    Memo              = "Auto-test deposit line"
                });
                miscIncomes.Add(inc);
            }

            List<MiscIncome> firstCompare  = new();
            List<MiscIncome> secondCompare = new();

            var addedDepositTxnIds  = new List<string>();
            var addedCustomerListIds = new List<string>();

            try
            {
                // ───────────────────────────────────────────────────────
                // 2) Push prerequisite customers, then first comparison
                // ───────────────────────────────────────────────────────
                using (var qb = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    foreach (var c in customers)
                    {
                        string listId = CustomerAdder.AddCustomer(qb, c);
                        Assert.False(string.IsNullOrWhiteSpace(listId), "Customer add failed");
                        addedCustomerListIds.Add(listId);

                        // If your deposit adder needs ReceivedFromListID, stash it:
                        foreach (var inc in miscIncomes)
                            inc.Lines.First(l => l.ReceivedFromName == c.Name)
                                 .ReceivedFromListID = listId;
                    }
                }

                // 2-b: first compare → expect all Added
                firstCompare = MiscIncomeComparator.CompareMiscIncomes(miscIncomes);

                foreach (var inc in firstCompare.Where(i => miscIncomes.Any(x => x.CompanyID == i.CompanyID)))
                {
                    Assert.Equal(MiscIncomeStatus.Added, inc.Status);
                    Assert.False(string.IsNullOrWhiteSpace(inc.TxnID));
                    addedDepositTxnIds.Add(inc.TxnID);
                }

                // ───────────────────────────────────────────────────────
                // 3) Mutate list to hit Different & Missing
                // ───────────────────────────────────────────────────────
                var mutated = new List<MiscIncome>(miscIncomes);
                var removed = mutated[0];                  // -> Missing
                var changed = mutated[1];                  // -> Different

                mutated.Remove(removed);
                changed.TotalAmount += 15.00m;

                secondCompare = MiscIncomeComparator.CompareMiscIncomes(mutated);
                var dict = secondCompare.ToDictionary(i => i.CompanyID);

                Assert.Equal(MiscIncomeStatus.Missing,   dict[removed.CompanyID].Status);
                Assert.Equal(MiscIncomeStatus.Different, dict[changed.CompanyID].Status);

                foreach (var stable in mutated.Where(m => m.CompanyID != changed.CompanyID))
                    Assert.Equal(MiscIncomeStatus.Unchanged, dict[stable.CompanyID].Status);
            }
            finally
            {
                // ───────────────────────────────────────────────────────
                // 4) Clean up: deposits first, then customers
                // ───────────────────────────────────────────────────────
                using (var qb = new QuickBooksSession(AppConfig.QB_APP_NAME))
                {
                    foreach (string txnId in addedDepositTxnIds)
                        DeleteDeposit(qb, txnId);

                    foreach (string listId in addedCustomerListIds)
                        DeleteCustomer(qb, listId);
                }
            }

            // ───────────────────────────────────────────────────────────
            // 5) Verify Serilog output
            // ───────────────────────────────────────────────────────────
            EnsureLogFileClosed();
            string logFile = GetLatestLogFile();
            EnsureLogFileExists(logFile);
            string text = File.ReadAllText(logFile);

            Assert.Contains("MiscIncomeComparator Initialized", text);
            Assert.Contains("MiscIncomeComparator Completed",   text);

            foreach (var inc in firstCompare.Concat(secondCompare))
            {
                string expected = $"MiscIncome {inc.CompanyID} is {inc.Status}.";
                Assert.Contains(expected, text);
            }
        }

        // ───────────────────────────────────────────────────────────
        // Helper - delete Deposit by TxnID
        // ───────────────────────────────────────────────────────────
        private void DeleteDeposit(QuickBooksSession qb, string txnId)
        {
            IMsgSetRequest rq = qb.CreateRequestSet();
            ITxnDel delRq    = rq.AppendTxnDelRq();
            delRq.TxnDelType.SetValue(ENTxnDelType.tdtDeposit);
            delRq.TxnID.SetValue(txnId);

            IMsgSetResponse rs = qb.SendRequest(rq);
            ShowResult(rs, "Deposit", txnId);
        }

        // Helper - delete Customer by ListID
        private void DeleteCustomer(QuickBooksSession qb, string listId)
        {
            IMsgSetRequest rq = qb.CreateRequestSet();
            IListDel delRq    = rq.AppendListDelRq();
            delRq.ListDelType.SetValue(ENListDelType.ldtCustomer);
            delRq.ListID.SetValue(listId);

            IMsgSetResponse rs = qb.SendRequest(rq);
            ShowResult(rs, "Customer", listId);
        }

        private void ShowResult(IMsgSetResponse rs, string kind, string id)
        {
            var rsp = rs.ResponseList?[0];
            Debug.WriteLine(rsp?.StatusCode == 0
                ? $"Deleted {kind} {id}"
                : $"Could not delete {kind} {id}: {rsp?.StatusMessage}");
        }
    }
}
