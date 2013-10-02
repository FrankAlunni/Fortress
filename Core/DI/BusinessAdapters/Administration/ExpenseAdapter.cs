// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExpenseAdapter.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the ExpenseAdapter type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using B1C.Utility.Helpers;

namespace B1C.SAP.DI.BusinessAdapters.Administration
{
    using System;
    using BusinessPartners;
    using Helpers;
    using SAPbobsCOM;

    /// <summary>
    /// Teh freight information
    /// </summary>
    public class ExpenseAdapter : SapObjectAdapter
    {
        #region Fields

        /// <summary>
        /// The Default Tax Code field
        /// </summary>
        private string defaultTaxCode;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="ExpenseAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="expnsCode">The expense code.</param>
        public ExpenseAdapter(Company company, int expnsCode)
            : base(company)
        {
            this.Expense = (AdditionalExpenses)this.Company.GetBusinessObject(BoObjectTypes.oAdditionalExpenses);

            if (!this.Expense.GetByKey(expnsCode))
            {
                throw new Exception(string.Format("Cound not fine the expense code{0}.", expnsCode));
            }
        }
        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the document.
        /// </summary>
        /// <value>The document.</value>
        public AdditionalExpenses Expense { get; protected set; }

        /// <summary>
        /// Gets the default tax code.
        /// </summary>
        /// <value>The default tax code.</value>
        public string DefaultTaxCode
        {
            get
            {
                if (string.IsNullOrEmpty(this.defaultTaxCode))
                {
                    ConfigurationHelper.GlobalConfiguration.Load(this.Company, "ExpensesTaxCode", out this.defaultTaxCode);
                }

                return this.defaultTaxCode;
            }
        }

        /// <summary>
        /// Releases the Documents COM Object.
        /// </summary>
        protected override void Release()
        {
            if (this.Expense != null)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(this.Expense);
                }
                finally
                {
                    this.Expense = null;
                    GC.Collect();
                }
            }
        }
        #endregion
    }
}