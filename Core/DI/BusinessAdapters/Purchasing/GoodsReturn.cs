//-----------------------------------------------------------------------
// <copyright file="GoodsReturn.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.Purchasing
{
    #region Using Directive(s)

    using System;
    using System.Collections.Generic;

    using Inventory;
    using SAPbobsCOM;
    using B1C.SAP.DI.Helpers;

    #endregion Using Directive(s)

    /// <summary>
    /// The good receipt adapter
    /// </summary>
    public class GoodsReturn : APDocumentAdapter
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="GoodsReturn"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="purchaseOrder">The purchase order.</param>
        public GoodsReturn(Company company, Documents purchaseOrder)
            : base(company, purchaseOrder)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GoodsReturn"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="docEntry">The doc entry.</param>
        public GoodsReturn(Company company, int docEntry)
            : base(company, docEntry)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GoodsReturn"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        public GoodsReturn(Company company)
            : base(company)
        {
            this.Document = (Documents)company.GetBusinessObject(BoObjectTypes.oPurchaseReturns);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GoodsReturn"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        public GoodsReturn(Company company, string xmlFile, int index)
            : base(company, xmlFile, index)
        {
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// The type of document (PurchaseOrder)
        /// </summary>
        public override BoObjectTypes DocumentType
        {
            get { return BoObjectTypes.oPurchaseReturns; }
        }

        /// <summary>
        /// Gets the target document type of this adapter's document type
        /// </summary>
        /// <value></value>
        public override BoObjectTypes TargetType
        {
            get { return BoObjectTypes.oPurchaseInvoices; }
        }

        #endregion Properties

        #region Method(s)
        #endregion Method(s)
    }
}
