//-----------------------------------------------------------------------
// <copyright file="APDocumentAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------


namespace B1C.SAP.DI.BusinessAdapters.Purchasing
{
    #region Using Directive(s)
    using System;
    using Administration;
    using Business;
    using SAPbobsCOM;
    using Utility.Logging;
    using B1C.SAP.DI.Helpers;

    #endregion Using Directive(s)

    /// <summary>
    /// The Payable Document Adapter
    /// </summary>
    public abstract class APDocumentAdapter : DocumentAdapter
    {
        #region Field(s)
        #endregion Field(s)

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="APDocumentAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        protected APDocumentAdapter(Company company)
            : base(company)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="APDocumentAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="docEntry">The id of the A/R Document</param>
        protected APDocumentAdapter(Company company, int docEntry)
            : base(company, docEntry)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="APDocumentAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="document">The A/R Document</param>
        protected APDocumentAdapter(Company company, Documents document)
            : base(company, document)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="APDocumentAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        protected APDocumentAdapter(Company company, string xmlFile, int index)
            : base(company, xmlFile, index)
        {
        }

        #endregion Constructor(s)

        #region Properties
        #endregion Properties

        #region Method(s)
        #endregion Method(s)
    }
}
