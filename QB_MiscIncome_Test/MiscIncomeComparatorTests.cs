using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Serilog;
using QBFC16Lib;
using QB_MiscIncome_Lib;          // MiscIncome, MiscIncomeLine, MiscIncomeStatus
using QB_MiscIncome_Lib;          // MiscIncomeComparator
using static QB_MiscIncome_Test.CommonMethods;

namespace QB_MiscIncome_Test
{
    [Collection("Sequential Tests")]
    public class MiscIncomeComparatorTests
    {
        [Fact]
        public void CompareMiscIncomes_InMemoryScenario_Verify_All_Statuses()
        {
            // ─── 1️⃣  Log prep ───────────────────────────────────────────────────
            EnsureLogFileClosed();
            DeleteOldLogFiles();
            ResetLogger();

            const int COMPANY_ID_START = 20_000;

            // ─── 2️⃣  Create QB fixtures (vendor + 2 items) ─────────────────────
            var createdVendorIds = new List<string>();
            var createdItemIds   = new List<string>();

            using (var qb = new QuickBooksSession(AppConfig.QB_APP_NAME))
            {
                string vendorListId = AddVendor(qb, $"TestVendor_{Guid.NewGuid():N}".Substring(0, 8));
                createdVendorIds.Add(vendorListId);

                for (int i = 0; i < 2; i++)
                {
                    string itemListId = AddInventoryItem(qb, $"TestPart_{Guid.NewGuid():N}".Substring(0, 8));
                    createdItemIds.Add(itemListId);
                }
            }

            // ─── 3️⃣  Build initial company misc-incomes (5 total) ───────────────
            var initialIncomes = new List<MiscIncome>();

            for (int i = 0; i < 4; i++)           // VALID ⇒ Added
                initialIncomes.Add(BuildValidIncome(i, COMPANY_ID_START,
                                                   createdVendorIds[0], createdItemIds));

            var invalidIncome = BuildInvalidIncome(COMPANY_ID_START + 4); // INVALID ⇒ FailedToAdd
            initialIncomes.Add(invalidIncome);

            List<MiscIncome> firstPass  = new();
            List<MiscIncome> secondPass = new();

            try
            {
                // ─── 4️⃣  First compare – expect Added & FailedToAdd ────────────
                firstPass = MiscIncomeComparator.CompareMiscIncomes(initialIncomes);

                foreach (var mi in firstPass.Where(m => m.InvoiceNum != invalidIncome.InvoiceNum))
                {
                    Assert.Equal(MiscIncomeStatus.Added, mi.Status);
                    Assert.False(string.IsNullOrEmpty(mi.TxnID));
                }

                var failed = firstPass.Single(m => m.InvoiceNum == invalidIncome.InvoiceNum);
                Assert.Equal(MiscIncomeStatus.FailedToAdd, failed.Status);
                Assert.True(string.IsNullOrEmpty(failed.TxnID));

                // ─── 5️⃣  Mutate list for second compare ───────────────────────
                var updatedIncomes = new List<MiscIncome>(initialIncomes);

                //     • Missing
                var incomeToRemove = updatedIncomes[0];
                updatedIncomes.Remove(incomeToRemove);

                //     • Different
                var incomeToModify = updatedIncomes[0];
                incomeToModify.Memo += "_MODIFIED";

                // ─── 6️⃣  Second compare – expect Missing / Different / … ──────
                secondPass = MiscIncomeComparator.CompareMiscIncomes(updatedIncomes);
                var secondDict = secondPass.ToDictionary(m => m.InvoiceNum);

                Assert.Equal(MiscIncomeStatus.Missing,   secondDict[incomeToRemove.InvoiceNum].Status);
                Assert.Equal(MiscIncomeStatus.Different, secondDict[incomeToModify.InvoiceNum].Status);

                foreach (var inv in updatedIncomes
                         .Where(m => m.InvoiceNum != incomeToModify.InvoiceNum &&
                                     m.InvoiceNum != invalidIncome.InvoiceNum)
                         .Select(m => m.InvoiceNum))
                {
                    Assert.Equal(MiscIncomeStatus.Unchanged, secondDict[inv].Status);
                }

                Assert.Equal(MiscIncomeStatus.FailedToAdd, secondDict[invalidIncome.InvoiceNum].Status);
            }
            finally
            {
                // ─── 7️⃣  QB clean-up (bills → items → vendor) ──────────────────
                using var qb = new QuickBooksSession(AppConfig.QB_APP_NAME);

                foreach (var mi in firstPass.Where(m => !string.IsNullOrEmpty(m.TxnID)))
                    DeleteBill(qb, mi.TxnID);

                foreach (var itemId in createdItemIds)
                    DeleteInventoryItem(qb, itemId);

                foreach (var vendorId in createdVendorIds)
                    DeleteVendor(qb, vendorId);
            }

            // ─── 8️⃣  Verify logs ──────────────────────────────────────────────
            EnsureLogFileClosed();
            string logFile = GetLatestLogFile();
            EnsureLogFileExists(logFile);
            string logs = File.ReadAllText(logFile);

            Assert.Contains("MiscIncomeComparator Initialized", logs);
            Assert.Contains("MiscIncomeComparator Completed",   logs);

            foreach (var mi in firstPass.Concat(secondPass))
                Assert.Contains($"MiscIncome {mi.InvoiceNum} is {mi.Status}.", logs);
        }

        // ────────────────────────── Helpers ───────────────────────────────────

        private MiscIncome BuildValidIncome(int idx, int companyStart,
                                            string vendorName, List<string> partNames) =>
            new()
            {
                VendorName = vendorName,
                BillDate   = DateTime.Today,
                InvoiceNum = $"INV_{Guid.NewGuid():N}".Substring(0, 10),
                Memo       = (companyStart + idx).ToString(),
                Lines = new()
                {
                    new MiscIncomeLine { PartName = partNames[0], Quantity = 2, UnitPrice = 15.5 },
                    new MiscIncomeLine { PartName = partNames[1], Quantity = 1, UnitPrice =  9.9 }
                }
            };

        private MiscIncome BuildInvalidIncome(int companyId) =>
            new()
            {
                VendorName = $"BadVendor_{Guid.NewGuid():N}".Substring(0, 6),
                BillDate   = DateTime.Today,
                InvoiceNum = $"INV_BAD_{Guid.NewGuid():N}".Substring(0, 10),
                Memo       = companyId.ToString(),
                Lines      = new() { new MiscIncomeLine { PartName = "BadItem", Quantity = 1, UnitPrice = 1.0 } }
            };

        // —— QuickBooks CRUD helpers (Vendor / Item / Bill) ————————————————
        private string AddVendor(QuickBooksSession s, string name)
        {
            var rq  = s.CreateRequestSet();
            var add = rq.AppendVendorAddRq();
            add.Name.SetValue(name);
            var rs  = s.SendRequest(rq).ResponseList.GetAt(0);
            if (rs.StatusCode != 0) throw new Exception(rs.StatusMessage);
            return ((IVendorRet)rs.Detail).ListID.GetValue();
        }

        private string AddInventoryItem(QuickBooksSession s, string name)
        {
            var rq  = s.CreateRequestSet();
            var add = rq.AppendItemInventoryAddRq();
            add.Name.SetValue(name);
            add.IncomeAccountRef.FullName.SetValue("Sales");
            add.COGSAccountRef.FullName.SetValue("Cost of Goods Sold");
            add.AssetAccountRef.FullName.SetValue("Inventory Asset");
            var rs  = s.SendRequest(rq).ResponseList.GetAt(0);
            if (rs.StatusCode != 0) throw new Exception(rs.StatusMessage);
            return ((IItemInventoryRet)rs.Detail).ListID.GetValue();
        }

        private void DeleteBill(QuickBooksSession s, string txnId)
        {
            var rq = s.CreateRequestSet();
            var del = rq.AppendTxnDelRq();
            del.TxnDelType.SetValue(ENTxnDelType.tdtBill);
            del.TxnID.SetValue(txnId);
            s.SendRequest(rq);
        }

        private void DeleteVendor(QuickBooksSession s, string listId) =>
            DeleteListObj(s, ENListDelType.ldtVendor, listId);

        private void DeleteInventoryItem(QuickBooksSession s, string listId) =>
            DeleteListObj(s, ENListDelType.ldtItemInventory, listId);

        private void DeleteListObj(QuickBooksSession s, ENListDelType type, string listId)
        {
            var rq  = s.CreateRequestSet();
            var del = rq.AppendListDelRq();
            del.ListDelType.SetValue(type);
            del.ListID.SetValue(listId);
            s.SendRequest(rq);
        }
    }
}
