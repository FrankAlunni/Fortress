// --------------------------------------------------------------------------------------------------------------------
// <copyright file="BusinessAdapterFactory.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the BusinessAdapterFactory type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters
{
    using BusinessPartners;
    using Constants;
    using Purchasing;
    using Sales;
    using SAPbobsCOM;

    /// <summary>
    /// BusinessAdapter Class Factory
    /// </summary>
    public class BusinessAdapterFactory
    {
        /// <summary>
        /// Creates the document adapter.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="docType">Type of the doc.</param>
        /// <param name="docEntry">The doc entry.</param>
        /// <returns>The document adapter requested.</returns>
        public static DocumentAdapter CreateDocumentAdapter(Company company, int docType, int docEntry)
        {
            DocumentAdapter adapter = null;

            BoObjectTypes documentType = SapDocumentTables.GetDocumentTypeFromCode(docType);

            switch (documentType)
            {
                case BoObjectTypes.oPurchaseReturns:
                    adapter = new GoodsReturn(company, docEntry);
                    break;
                case BoObjectTypes.oPurchaseInvoices:
                    adapter = new APInvoiceAdapter(company, docEntry);
                    break;
                case BoObjectTypes.oPurchaseDeliveryNotes:
                    adapter = new GoodsReceiptAdapter(company, docEntry);
                    break;
                case BoObjectTypes.oPurchaseOrders:
                    adapter = new PurchaseOrderAdapter(company, docEntry);
                    break;
                case BoObjectTypes.oInvoices:
                    adapter = new ARInvoiceAdapter(company, docEntry);
                    break;
                case BoObjectTypes.oDeliveryNotes:
                    adapter = new DeliveryAdapter(company, docEntry);
                    break;
                case BoObjectTypes.oOrders:
                    adapter = new SalesOrderAdapter(company, docEntry);
                    break;
            }

            return adapter;
        }

        /// <summary>
        /// Creates a DocumentAdapter instance of the provided Document type, or null if none exists
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="document">The document to create an adapter around</param>
        /// <returns>A DocumentAdapter</returns>
        public static DocumentAdapter CreateDocumentAdapter(Company company, Documents document)
        {
            DocumentAdapter adapter = null;
            BoObjectTypes documentType = document.DocObjectCode;

            switch (documentType)
            {
                case BoObjectTypes.oPurchaseReturns:
                    adapter = new GoodsReturn(company, document);
                    break;
                case BoObjectTypes.oPurchaseInvoices:
                    adapter = new APInvoiceAdapter(company, document);
                    break;
                case BoObjectTypes.oPurchaseDeliveryNotes:
                    adapter = new GoodsReceiptAdapter(company, document);
                    break;
                case BoObjectTypes.oPurchaseOrders:
                    adapter = new PurchaseOrderAdapter(company, document);
                    break;
                case BoObjectTypes.oInvoices:
                    adapter = new ARInvoiceAdapter(company, document);
                    break;
                case BoObjectTypes.oDeliveryNotes:
                    adapter = new DeliveryAdapter(company, document);
                    break;
                case BoObjectTypes.oOrders:
                    adapter = new SalesOrderAdapter(company, document);
                    break;
                default:
                    adapter = new DraftAdapter(company, document);
                    break;
            }

            return adapter;
        }

        /// <summary>
        /// Creates the document adapter.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="docType">Type of the doc.</param>
        /// <returns>A DocumentAdapter</returns>
        public static DocumentAdapter CreateDocumentAdapter(Company company, int docType)
        {
            DocumentAdapter adapter = null;

            BoObjectTypes documentType = SapDocumentTables.GetDocumentTypeFromCode(docType);

            switch (documentType)
            {
                case BoObjectTypes.oPurchaseReturns:
                    adapter = new GoodsReturn(company);
                    break;
                case BoObjectTypes.oPurchaseInvoices:
                    adapter = new APInvoiceAdapter(company);
                    break;
                case BoObjectTypes.oPurchaseDeliveryNotes:
                    adapter = new GoodsReceiptAdapter(company);
                    break;
                case BoObjectTypes.oPurchaseOrders:
                    adapter = new PurchaseOrderAdapter(company);
                    break;
                case BoObjectTypes.oInvoices:
                    adapter = new ARInvoiceAdapter(company);
                    break;
                case BoObjectTypes.oDeliveryNotes:
                    adapter = new DeliveryAdapter(company);
                    break;
                case BoObjectTypes.oOrders:
                    adapter = new SalesOrderAdapter(company);
                    break;
            }

            return adapter;
        }

        /// <summary>
        /// Creates the document adapter.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="documentType">Type of the document.</param>
        /// <returns>A DocumentAdapter</returns>
        public static DocumentAdapter CreateDocumentAdapter(Company company, BoObjectTypes documentType)
        {
            DocumentAdapter adapter = null;

            switch (documentType)
            {
                case BoObjectTypes.oPurchaseReturns:
                    adapter = new GoodsReturn(company);
                    break;
                case BoObjectTypes.oPurchaseInvoices:
                    adapter = new APInvoiceAdapter(company);
                    break;
                case BoObjectTypes.oPurchaseDeliveryNotes:
                    adapter = new GoodsReceiptAdapter(company);
                    break;
                case BoObjectTypes.oPurchaseOrders:
                    adapter = new PurchaseOrderAdapter(company);
                    break;
                case BoObjectTypes.oInvoices:
                    adapter = new ARInvoiceAdapter(company);
                    break;
                case BoObjectTypes.oDeliveryNotes:
                    adapter = new DeliveryAdapter(company);
                    break;
                case BoObjectTypes.oOrders:
                    adapter = new SalesOrderAdapter(company);
                    break;
            }

            return adapter;
        }
    }
}
