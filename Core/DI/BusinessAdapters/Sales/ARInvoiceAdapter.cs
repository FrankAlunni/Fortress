//-----------------------------------------------------------------------
// <copyright file="ARInvoiceAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------
using B1C.Utility.Helpers;

namespace B1C.SAP.DI.BusinessAdapters.Sales
{
    #region Using Directive(s)
    using System;
    using System.Collections.Generic;
    using System.Data;
    using Business;
    using Components;
    using Helpers;
    using SAPbobsCOM;
    using Utility.Logging;
    using B1C.SAP.DI.Model.Exceptions;

    #endregion Using Directive(s)

    /// <summary>
    /// The Receipvable Invoice Adapter
    /// </summary>
    public class ARInvoiceAdapter : ARDocumentAdapter
    {
        #region Private Static Fields

        /// <summary>
        /// The Lock object
        /// </summary>
        private static readonly object LockingObject = new object();

        
        #endregion Private Static Fields

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="ARInvoiceAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="receivableInvoice">The receivable invoice.</param>
        public ARInvoiceAdapter(Company company, Documents receivableInvoice) : base(company, receivableInvoice)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ARInvoiceAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="docEntry">The existing invoice document's DocEntry</param>
        public ARInvoiceAdapter(Company company, int docEntry) : base(company, docEntry)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ARInvoiceAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        public ARInvoiceAdapter(Company company) : base(company)
        {
            this.Document = (Documents)company.GetBusinessObject(BoObjectTypes.oInvoices);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ARInvoiceAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        public ARInvoiceAdapter(Company company, string xmlFile, int index)
            : base(company, xmlFile, index)
        { 
        }

        #endregion Constructor(s)

        #region Properties
        /// <summary>
        /// The type of document (PurchaseOrder)
        /// </summary>
        public override BoObjectTypes DocumentType
        {
            get { return BoObjectTypes.oInvoices; }
        }

        /// <summary>
        /// The target document (CreditNote)
        /// </summary>
        public override BoObjectTypes TargetType
        {
            get { return BoObjectTypes.oCreditNotes; }
        }

        #endregion Properties

        #region Method(s)
        /// <summary>
        /// Determines whether [is delivery partially invoiced].
        /// </summary>
        /// <returns>
        /// Returns true if the delivery has already been invocied
        /// </returns>
        private bool IsBaseDocumentPartiallyInvoiced()
        {
            var invoiceLines = this.Document.Lines;
            try
            {
                for (int invoiceRowIndex = 0; invoiceRowIndex < invoiceLines.Count; invoiceRowIndex++)
                {
                    invoiceLines.SetCurrentLine(invoiceRowIndex);

                    if (this.IsBaseSalesOrderPartiallyInvoiced(invoiceLines))
                    {
                        return true;
                    }
                }
            }
            finally
            {
                COMHelper.Release(ref invoiceLines);
            }

            return false;
        }

        /// <summary>
        /// Determines whether [is delivery line partially invoiced] [the specified current line].
        /// </summary>
        /// <param name="currentLine">The current line.</param>
        /// <returns>
        /// <c>true</c> if [is delivery line partially invoiced] [the specified current line]; otherwise, <c>false</c>.
        /// </returns>
        private bool IsBaseLinePartiallyInvoiced(IDocument_Lines currentLine)
        {
            var query = new SqlHelper();
            query.Builder.AppendLine("SELECT * FROM DLN1 T0");
            query.Builder.AppendLine("JOIN RDR1 T1 ON T0.BaseEntry = T1.DocEntry AND T0.BaseLine = T1.LineNum");
            query.Builder.AppendLine("JOIN DLN1 T2 ON T2.BaseEntry = T1.DocEntry AND T2.BaseLine = T1.LineNum");
            query.Builder.AppendFormat(" WHERE T0.DocEntry = {0} AND T0.LineNum = {1} AND T0.TrgetEntry IS NOT NULL", currentLine.BaseEntry, currentLine.BaseLine);
            using (var rs = new RecordsetAdapter(this.Company, query.ToString()))
            {
                return !rs.EoF;
            }
        }

        /// <summary>
        /// Determines whether [is base sales order partially invoiced] [the specified current line].
        /// </summary>
        /// <param name="currentLine">The current line.</param>
        /// <returns>
        /// <c>true</c> if [is base sales order partially invoiced] [the specified current line]; otherwise, <c>false</c>.
        /// </returns>
        private bool IsBaseSalesOrderPartiallyInvoiced(IDocument_Lines currentLine)
        {
            var query = new SqlHelper();
            query.Builder.AppendLine("SELECT * FROM DLN1 T0");
            query.Builder.AppendLine("JOIN RDR1 T1 ON T0.BaseEntry = T1.DocEntry AND T0.BaseLine = T1.LineNum");
            query.Builder.AppendLine("JOIN DLN1 T2 ON T2.BaseEntry = T1.DocEntry");
            query.Builder.AppendFormat(" WHERE T0.DocEntry = {0} AND T0.LineNum = {1} AND T0.TrgetEntry IS NOT NULL", currentLine.BaseEntry, currentLine.BaseLine);
            using (var rs = new RecordsetAdapter(this.Company, query.ToString()))
            {
                return !rs.EoF;
            }
        }

        #endregion Method(s)
    }
}
