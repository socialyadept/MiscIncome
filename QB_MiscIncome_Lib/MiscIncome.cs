using System;
using System.Collections.Generic;

namespace QB_MiscIncome_Lib
{
    /// <summary>
    /// Represents a deposit transaction from QuickBooks with multiple lines
    /// </summary>
    public class MiscIncome
    {
        /// <summary>
        /// The QuickBooks transaction ID
        /// </summary>
        public string TxnID { get; set; } = "";

        /// <summary>
        /// The date of the deposit
        /// </summary>
        public DateTime DepositDate { get; set; }

        /// <summary>
        /// The memo field of the deposit, used to store CompanyID
        /// </summary>
        public string Memo { get; set; } = "";

        /// <summary>
        /// The account the deposit was made to
        /// </summary>
        public string DepositToAccount { get; set; } = "";

        /// <summary>
        /// The total amount of the deposit
        /// </summary>
        public double TotalAmount { get; set; }

        /// <summary>
        /// The list of deposit lines
        /// </summary>
        public List<MiscIncomeLine> Lines { get; set; } = new List<MiscIncomeLine>();
    }

    /// <summary>
    /// Represents a line item in a deposit transaction
    /// </summary>
    public class MiscIncomeLine
    {
        /// <summary>
        /// The customer or entity that the money was received from
        /// </summary>
        public string ReceivedFromName { get; set; } = "";

        /// <summary>
        /// The ListID of the customer or entity
        /// </summary>
        public string ReceivedFromListID { get; set; } = "";

        /// <summary>
        /// The account the money is attributed to (e.g., "Sales", "Shipping and Delivery Income")
        /// </summary>
        public string FromAccountName { get; set; } = "";

        /// <summary>
        /// The ListID of the account
        /// </summary>
        public string FromAccountListID { get; set; } = "";

        /// <summary>
        /// The amount for this line
        /// </summary>
        public double Amount { get; set; }

        /// <summary>
        /// Optional memo for this specific line
        /// </summary>
        public string Memo { get; set; } = "";
    }
}