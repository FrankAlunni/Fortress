//-----------------------------------------------------------------------
// <copyright file="GoodsReceiptAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.Purchasing
{
    #region Using Directive(s)

    using System;
    using System.Collections.Generic;

    using Helpers;
    using Inventory;
    using SAPbobsCOM;

    #endregion Using Directive(s)

    /// <summary>
    /// The good receipt adapter
    /// </summary>
    public class GoodsReceiptAdapter : APDocumentAdapter
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="GoodsReceiptAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="document">The base document.</param>
        public GoodsReceiptAdapter(Company company, Documents document)
            : base(company, document)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GoodsReceiptAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="docEntry">The doc entry.</param>
        public GoodsReceiptAdapter(Company company, int docEntry)
            : base(company, docEntry)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GoodsReceiptAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        public GoodsReceiptAdapter(Company company)
            : base(company)
        {
            this.Document = (Documents)company.GetBusinessObject(BoObjectTypes.oPurchaseDeliveryNotes);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GoodsReceiptAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        public GoodsReceiptAdapter(Company company, string xmlFile, int index)
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
            get { return BoObjectTypes.oPurchaseDeliveryNotes; }
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
        /// <summary>
        /// Determines whether [is sub vendor invoice].
        /// </summary>
        /// <returns>
        /// <c>true</c> if [is sub vendor invoice]; otherwise, <c>false</c>.
        /// </returns>
        public bool IsSubVendorGoodReceipt()
        {
            return !string.IsNullOrEmpty(this.Document.Reference2);
        }

        /// <summary>
        /// The base Good receipts.
        /// </summary>
        /// <returns>A list of base good receipt POs</returns>
        public GoodsReceiptAdapter[] BaseGoodReceiptPOs()
        {
            if (this.IsSubVendorGoodReceipt())
            {
                try
                {
                    int baseEntry;
                    int.TryParse(this.Document.Reference2, out baseEntry);

                    if (baseEntry != 0)
                    {
                        var baseDoc = (Documents)this.Company.GetBusinessObject(BoObjectTypes.oPurchaseDeliveryNotes);
                        if (baseDoc.GetByKey(baseEntry))
                        {
                            return new[] { new GoodsReceiptAdapter(this.Company, baseDoc) };
                        }
                    }

                    return null;
                }
                catch
                {
                    return null;
                }
            }

            var adapters = new Dictionary<int, GoodsReceiptAdapter>();
            int currentLine = this.GetLineIndex(this.Document.Lines.LineNum);

            for (int lineIndex = 0; lineIndex < this.Document.Lines.Count; lineIndex++)
            {
                this.Document.Lines.SetCurrentLine(lineIndex);
                var document = this.GetLineBaseDocument(this.Document.Lines.LineNum);
                if (!adapters.ContainsKey(document.DocEntry))
                {
                    adapters.Add(document.DocEntry, new GoodsReceiptAdapter(this.Company, document));
                }
            }

            if (adapters.Count > 0)
            {
                var results = new GoodsReceiptAdapter[adapters.Count];
                adapters.Values.CopyTo(results, 0);

                this.Document.Lines.SetCurrentLine(currentLine);
                return results;
            }

            return null;
        }

        /// <summary>
        /// Creates an A/P Invoice from this DeliveryNote
        /// </summary>
        /// <returns>The newly created (unsaved) A/P Invoice</returns>
        /// <remarks>
        /// The returned document is not saved to the DB. It is the responsibility of the calling
        /// object to call the Add() method on the object if it is wished to be saved.
        /// </remarks>
        public APInvoiceAdapter CopyToAPInvoice()
        {
            Documents payableInvoice = CopyAllToDocument(BoObjectTypes.oPurchaseInvoices);
            return new APInvoiceAdapter(this.Company, payableInvoice);
        }

        /// <summary>
        /// Copies the lines to AP invoice.
        /// </summary>
        /// <param name="lines">The lines to copy.</param>
        /// <returns>The newly created (unsaved) A/P Invoice</returns>
        /// <remarks>
        /// The returned document is not saved to the DB. It is the responsibility of the calling
        /// object to call the Add() method on the object if it is wished to be saved.
        /// </remarks>
        public APInvoiceAdapter CopyLinesToAPInvoice(int[] lines)
        {
            Documents payableInvoice = CopyLinesToDocument(BoObjectTypes.oPurchaseInvoices, lines);
            return new APInvoiceAdapter(this.Company, payableInvoice);
        }

        /// <summary>
        /// Determines whether [is from vendor invoice].
        /// </summary>
        /// <returns>
        /// <c>true</c> if [is from vendor invoice]; otherwise, <c>false</c>.
        /// </returns>
        public bool IsFromVendorInvoice()
        {
            // Check if this GR is coming from an A/P Invoice file 
            // from Vendor (it will have a reference # if so)
            return !string.IsNullOrEmpty(this.Document.Reference2);
        }

        /// <summary>
        /// Determines whether [is from reserve invoice].
        /// </summary>
        /// <returns>
        /// <c>true</c> if [is from reserve invoice]; otherwise, <c>false</c>.
        /// </returns>
        public bool IsFromReserveInvoice()
        {
            bool isFromReserve = false;
            Document_Lines line = this.Document.Lines;

            try
            {
                // If the base document type is an invoice (not a PO), 
                // then this is based on a reserve invoice
                if (line.BaseType == (int)BoObjectTypes.oPurchaseInvoices)
                {
                    isFromReserve = true;
                }
            }
            finally
            {
                COMHelper.Release(ref line);
            }

            return isFromReserve;
        }

        /// <summary>
        /// Copies the quantities to return.
        /// </summary>
        /// <param name="itemQuantities">The item quantities.</param>
        /// <returns>The return document</returns>
        public GoodsReturn CopyQuantitiesToReturn(IDictionary<string, int> itemQuantities)
        {
            IDictionary<int, int> quantitiesToReceivePerLine = new Dictionary<int, int>();
            IDictionary<string, int> quantitiesHandled = new Dictionary<string, int>();

            // Initialize the handled quantities to 0 for each item
            foreach (string itemCode in itemQuantities.Keys)
            {
                quantitiesHandled[itemCode] = 0;
            }

            // Determine the base lines to pull forward
            IDictionary<int, int> targetQuantityByLine = this.GetLineTargetQuantities();

            foreach (Document_Lines line in this.Lines)
            {
                // If the line is closed, we can't copy to target
                if (line.Quantity <= targetQuantityByLine[line.LineNum])
                {
                    continue;
                }

                // Check if this item code is needed to be copied
                if (itemQuantities.ContainsKey(line.ItemCode))
                {
                    // Check how many of these items we still need to copy forward
                    int remainingQuantity = itemQuantities[line.ItemCode] - quantitiesHandled[line.ItemCode];

                    if (remainingQuantity > 0)
                    {
                        int quantityToHandle = Math.Min(remainingQuantity, (int)line.Quantity);

                        quantitiesToReceivePerLine[line.LineNum] = quantityToHandle;
                        quantitiesHandled[line.ItemCode] += quantityToHandle;
                    }
                }
            }

            var targetDoc = BusinessAdapterFactory.CreateDocumentAdapter(this.Company, BoObjectTypes.oPurchaseReturns);
            
            // Fill out the document level info            
            targetDoc.Document.CardCode = this.Document.CardCode;
            targetDoc.Document.SalesPersonCode = this.Document.SalesPersonCode;
            targetDoc.Document.DocType = BoDocumentTypes.dDocument_Items;
            targetDoc.Document.DocumentSubType = BoDocumentSubType.bod_None;
            targetDoc.Document.Address = this.Document.Address;
            targetDoc.Document.ShipToCode = this.Document.ShipToCode;
            targetDoc.Document.Address2 = this.Document.Address2;
            targetDoc.Document.HandWritten = BoYesNoEnum.tNO;
            targetDoc.Document.JournalMemo = string.Format("{0} - {1}.", "Return", this.Document.CardCode);
            targetDoc.SetComments(string.Format("Based on {0} {1}.", "Good Receipt PO", this.Document.DocEntry));
            targetDoc.Document.PaymentGroupCode = this.Document.PaymentGroupCode;
            targetDoc.Document.DocDueDate = DateTime.Today;

            // Transfer all the custom fields value
            this.TransferFieldsToAdapter(targetDoc);

            // Now fill out the lines in the target
            var targetLine = targetDoc.Document.Lines;
            var baseLine = this.Document.Lines;
            try
            {
                int count = 0;
                foreach (int lineNum in quantitiesToReceivePerLine.Keys)
                {
                    baseLine.SetCurrentLine(count);
                    int lineQty = quantitiesToReceivePerLine[lineNum];

                    targetLine.BaseLine = lineNum;
                    targetLine.BaseEntry = this.Document.DocEntry;
                    targetLine.BaseType = (int)this.Document.DocObjectCode;
                    targetLine.Quantity = lineQty;
                    targetLine.ItemCode = baseLine.ItemCode;
                    targetLine.TaxCode = baseLine.TaxCode;
                    targetLine.Rate = baseLine.Rate;

                    this.TransferLineFieldsToAdapter(targetDoc);    
                    this.TransferLineExpensesToAdapter(targetDoc);

                    count++;
                    targetLine.Add();
                }
                targetLine.SetCurrentLine(0);
            }
            finally
            {
                COMHelper.Release(ref baseLine);
                COMHelper.Release(ref targetLine);
            }

            return (GoodsReturn)targetDoc;
        }
        #endregion Method(s)
    }
}
