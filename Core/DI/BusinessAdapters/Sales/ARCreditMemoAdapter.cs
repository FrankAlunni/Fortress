//-----------------------------------------------------------------------
// <copyright file="ARCreditMemoAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.Sales
{
    #region Using Directive(s)
    using SAPbobsCOM;

    #endregion Using Directive(s)

    /// <summary>
    /// The Payable Invoice Adapter
    /// </summary>
    public class ARCreditMemoAdapter : ARDocumentAdapter
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="ARCreditMemoAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="baseDocument">The base document.</param>
        public ARCreditMemoAdapter(Company company, Documents baseDocument)
            : base(company, baseDocument)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ARCreditMemoAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="docEntry">The doc entry.</param>
        public ARCreditMemoAdapter(Company company, int docEntry)
            : base(company, docEntry)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ARCreditMemoAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        public ARCreditMemoAdapter(Company company, string xmlFile, int index)
            : base(company, xmlFile, index)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ARCreditMemoAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        public ARCreditMemoAdapter(Company company)
            : base(company)
        {
            this.Document = (Documents)company.GetBusinessObject(BoObjectTypes.oCreditNotes);
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// The type of document (A/R Credit Note)
        /// </summary>
        public override BoObjectTypes DocumentType
        {
            get { return BoObjectTypes.oCreditNotes; }
        }

        /// <summary>
        /// The target document (Return)
        /// </summary>
        public override BoObjectTypes TargetType
        {
            get { return BoObjectTypes.oReturns; }
        }

        #endregion Properties

        #region Method(s)

        #endregion Method(s)
    }
}
