using System;
using System.Collections.Generic;

namespace QB_MiscIncome_Lib
{
    /// Represents a deposit transaction from QuickBooks with multiple lines
    public class MiscIncome
    {
        /// The QuickBooks transaction ID
        public string TxnID { get; set; } = "";

        /// The date of the deposit
        public DateTime DepositDate { get; set; }

        /// The memo field of the deposit, used to store CompanyID
        public string Memo { get; set; } = "";

        /// The account the deposit was made to
        public string DepositToAccount { get; set; } = "";

        /// The total amount of the deposit
        public double TotalAmount { get; set; }

        /// The list of deposit lines
        public List<MiscIncomeLine> Lines { get; set; } = new List<MiscIncomeLine>();
    }

    /// <summary>
    /// Represents a line item in a deposit transaction
    /// </summary>
    public class MiscIncomeLine
    {
        /// The customer or entity that the money was received from

        public string ReceivedFromName { get; set; } = "";

        /// The ListID of the customer or entity
        public string ReceivedFromListID { get; set; } = "";

        /// The account the money is attributed to (e.g., "Sales", "Shipping and Delivery Income")
        public string FromAccountName { get; set; } = "";

        /// The ListID of the account
        public string FromAccountListID { get; set; } = "";

        /// The amount for this line
        public double Amount { get; set; }

        /// Optional memo for this specific line
        public string Memo { get; set; } = "";
    }
}