//-----------------------------------------------------------------------
// <copyright file="ARDocumentAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------
using B1C.SAP.DI.Helpers;

namespace B1C.SAP.DI.BusinessAdapters.Sales
{
    #region Using Directive(s)

    using System.Collections.Generic;
    using System.Text;

    using SAPbobsCOM;

    #endregion Using Directive(s)

    /// <summary>
    /// The Receivable Document Adapter
    /// </summary>
    public abstract class ARDocumentAdapter : DocumentAdapter
    {
        #region Field(s)

        /// <summary>
        /// The id of the base sales order for this document
        /// </summary>
        private IDictionary<int, int> salesOrderDocEntries;

        #endregion Field(s)

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="ARDocumentAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        protected ARDocumentAdapter(Company company)
            : base(company)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ARDocumentAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="docEntry">The id of the A/R Document</param>
        protected ARDocumentAdapter(Company company, int docEntry)
            : base(company, docEntry)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ARDocumentAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="document">The A/R Document</param>
        protected ARDocumentAdapter(Company company, Documents document)
            : base(company, document)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ARDocumentAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        protected ARDocumentAdapter(Company company, string xmlFile, int index)
            : base(company, xmlFile, index)
        {
        }

        #endregion Constructor(s)

        #region Properties

        /// <summary>
        /// Gets the list of all base Sales Order DocEntry values, mapped from the line number of this document
        /// </summary>
        public IDictionary<int, int> SalesOrderDocEntries
        {
            get
            {
                if (this.salesOrderDocEntries == null)
                {
                    this.salesOrderDocEntries = this.GetBaseSalesOrderIds();
                }

                return this.salesOrderDocEntries;
            }
        }

        #endregion Properties

        #region Method(s)
        /// <summary>
        /// Helper method that gets the DocEntry of the base sales orders that originated this document
        /// (or the DocEntry of this document, if this is a Sales Order)
        /// </summary>
        /// <returns>A dictionary mapping lines in the document to sales orders</returns>
        private IDictionary<int, int> GetBaseSalesOrderIds()
        {            
            IDictionary<int, int> orderIds = new Dictionary<int, int>();
            using(var lines = new Lines(this.Document.Lines))
            {
                foreach (Document_Lines line in lines)
                {
                    orderIds.Add(line.LineNum, this.GetAssociatedOrderId(line));
                }
            }

            return orderIds;
        }

        /// <summary>
        /// Helper method to retrieve the base sales order id of the given document line
        /// </summary>
        /// <param name="line">The document line</param>
        /// <returns>
        /// The associated sales order id of the given document line
        /// </returns>
        /// <remarks>This method will return -1 if the base line cannot be found</remarks>
        private int GetAssociatedOrderId(IDocument_Lines line)
        {
            // If this is a Sales Order document, it's always *this* order id
            if (this.DocumentType == BoObjectTypes.oOrders)
            {
                return this.Document.DocEntry;
            }
            
            // Recursively grab the line from the next parent, until we get to the Sales Order, or we run out of parents
            if (line.BaseType == -1 || (BoAPARDocumentTypes)line.BaseType == BoAPARDocumentTypes.bodt_Order)
            {
                return line.BaseEntry;
            }

            // Grab the base document
            var document = (Documents)this.Company.GetBusinessObject((BoObjectTypes)line.BaseType);
            document.GetByKey(line.BaseEntry);
            Document_Lines subLine = document.Lines;
            subLine.SetCurrentLine(line.BaseLine);

            // Make the recursive call
            try
            {
                return this.GetAssociatedOrderId(subLine);
            }
            finally
            {
                COMHelper.Release(ref subLine);
            }
        }

        #endregion Method(s)
    }
}
