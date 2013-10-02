//-----------------------------------------------------------------------
// <copyright file="APCreditMemoAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.Purchasing
{
    #region Using Directive(s)
    using SAPbobsCOM;

    #endregion Using Directive(s)

    /// <summary>
    /// The Payable Invoice Adapter
    /// </summary>
    public class APCreditMemoAdapter : APDocumentAdapter
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="APCreditMemoAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="baseDocument">The base document.</param>
        public APCreditMemoAdapter(Company company, Documents baseDocument)
            : base(company, baseDocument)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="APCreditMemoAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="docEntry">The doc entry.</param>
        public APCreditMemoAdapter(Company company, int docEntry)
            : base(company, docEntry)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="APCreditMemoAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        public APCreditMemoAdapter(Company company, string xmlFile, int index)
            : base(company, xmlFile, index)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="APCreditMemoAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        public APCreditMemoAdapter(Company company)
            : base(company)
        {
            this.Document = (Documents)company.GetBusinessObject(BoObjectTypes.oPurchaseCreditNotes);
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// The type of document (A/P Credit Memo)
        /// </summary>
        public override BoObjectTypes DocumentType
        {
            get { return BoObjectTypes.oPurchaseCreditNotes; }
        }

        /// <summary>
        /// The target document (Goods Return)
        /// </summary>
        public override BoObjectTypes TargetType
        {
            get { return BoObjectTypes.oPurchaseReturns; }
        }

        #endregion Properties

        #region Method(s)

        #endregion Method(s)
    }
}
