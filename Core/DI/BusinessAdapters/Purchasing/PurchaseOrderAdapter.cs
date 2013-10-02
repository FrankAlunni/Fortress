//-----------------------------------------------------------------------
// <copyright file="PurchaseOrderAdapter.cs" company="B1C Canada Inc.">
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
    using Sales;
    using SAPbobsCOM;
    using Utility.Helpers;
    using Utility.Logging;
    #endregion Using Directive(s)

    /// <summary>
    /// The Purchase Order Adapter
    /// </summary>
    public class PurchaseOrderAdapter : APDocumentAdapter
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="PurchaseOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="purchaseOrder">The purchase order.</param>
        public PurchaseOrderAdapter(Company company, Documents purchaseOrder) : base(company, purchaseOrder)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PurchaseOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="docEntry">The doc entry.</param>
        public PurchaseOrderAdapter(Company company, int docEntry) : base(company, docEntry)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PurchaseOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        public PurchaseOrderAdapter(Company company)
            : base(company)
        {
            this.Document = (Documents)company.GetBusinessObject(BoObjectTypes.oPurchaseOrders);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PurchaseOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        public PurchaseOrderAdapter(Company company, string xmlFile, int index)
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
            get { return BoObjectTypes.oPurchaseOrders; }
        }

        /// <summary>
        /// Gets the target document type of this adapter's document type
        /// </summary>
        /// <value>The Object Target Type</value>
        public override BoObjectTypes TargetType
        {
            get { return BoObjectTypes.oPurchaseDeliveryNotes; }
        }

        /// <summary>
        /// Gets or sets the drop ship sales order id.
        /// </summary>
        /// <value>The drop ship sales order id.</value>
        public int DropShipSalesOrderId
        {
            get; set;
        }

        /// <summary>
        /// Gets or sets the drop ship sales order line id.
        /// </summary>
        /// <value>The drop ship sales order line id.</value>
        public int DropShipSalesOrderLineId
        {
            get; set;
        }
        #endregion Properties

        #region Methods
        /// <summary>
        /// Creates a Goods Receipt PO based on this PO
        /// </summary>
        /// <remarks>
        /// The returned document is not saved to the DB. It is the responsibility of the calling
        /// object to call the Add() method on the object if it is wished to be saved.
        /// </remarks>
        /// <returns>The newly created (unsaved) Goods Receipt PO</returns>
        public Documents CopyToGoodsReceipt()
        {
            var goodsReceipt = CopyAllToDocument(BoObjectTypes.oPurchaseDeliveryNotes);
            return goodsReceipt;
        }

        /// <summary>
        /// Gets all invoiced quantities for this purchase order
        /// </summary>
        /// <returns>A list of invoiced quantities</returns>
        public IDictionary<string, int> GetInvoicedQuantities()
        {
            IDictionary<string, int> invoicedQuantities = new Dictionary<string, int>();

            // Begin by assuming that there are no invoices
            foreach (Document_Lines line in this.Lines)
            {
                invoicedQuantities[line.ItemCode] = 0;
            }

            IList<Documents> receipts = this.GetAllTargetDocuments();
            foreach (Documents receipt in receipts)
            {
                var goodsReceipt = new GoodsReceiptAdapter(this.Company, receipt);
                try
                {
                    Lines receiptLines = goodsReceipt.Lines;

                    foreach (Document_Lines line in receiptLines)
                    {
                        if (line.BaseEntry == this.Document.DocEntry)
                        {
                            // This GR line comes from this PO -> Get the Invoices for this line
                            IList<Documents> invoices = goodsReceipt.GetTargetDocuments(line.LineNum);
                            foreach (Documents invoice in invoices)
                            {
                                try
                                {
                                    using (var invoiceLines = new Lines(invoice.Lines))
                                    {
                                        foreach (Document_Lines invoiceLine in invoiceLines)
                                        {
                                            if (invoiceLine.BaseEntry == receipt.DocEntry && invoiceLine.BaseLine == line.LineNum)
                                            {
                                                invoicedQuantities[line.ItemCode] += (int)line.Quantity;
                                                break;
                                            }
                                        }
                                    }
                                }
                                finally
                                {
                                    COMHelper.ReleaseOnly(invoice);
                                }
                            }
                        }
                    }
                }
                finally
                {
                    goodsReceipt.Dispose();   
                }

                COMHelper.ReleaseOnly(receipt);
            }

            return invoicedQuantities;
        }

        /// <summary>
        /// Updates the ship via.
        /// </summary>
        public void UpdateShipVia()
        {
            var order = this.DropShipSalesOrder();
            if (order != null)
            {
                try
                {
                    this.Document.TransportationCode = this.GetTransportationCode();

                    if (this.DropShipSalesOrderId > 0)
                    {
                        var sql = new SqlHelper();
                        sql.Builder.AppendLine("SELECT TrnspCode");
                        sql.Builder.AppendLine("FROM [@XX_PPCP] T0");
                        sql.Builder.AppendLine("JOIN OSHP T1 ON T0.U_ShipRqOt = T1.TrnspName");
                        sql.Builder.AppendLine("WHERE U_ShipReq = 'O' and U_CardCode = @CardCode");
                        sql.AddParameter("@CardCode", System.Data.DbType.String, order.Document.CardCode);

                        using (var shipVia = new RecordsetAdapter(this.Company, sql.ToString()))
                        {
                            if (!shipVia.EoF)
                            {
                                this.Document.TransportationCode = Convert.ToInt32(shipVia.FieldValue("TrnspCode").ToString());
                            }
                        }
                    }
                }
                finally
                {
                    order.Dispose();   
                }
            }
        }

        /// <summary>
        /// Drops the ship sales order.
        /// </summary>
        /// <returns>Returns an instance of the droppedshipped sales order adapter</returns>
        public SalesOrderAdapter DropShipSalesOrder()
        {
            string task = "Retrieving the sales order from the purchase order ";
            ThreadedAppLog.WriteLine("    " + task);
            try
            {
                this.DropShipSalesOrderId = Convert.ToInt32(this.GetUserDefinedField("U_SrcDoc").ToString());
                this.DropShipSalesOrderLineId = Convert.ToInt32(this.GetUserDefinedField("U_SrcDocLn").ToString());
            }
            catch
            {
                return null;
            }

            if (this.DropShipSalesOrderId > 0)
            {
                task = string.Format(
                    "Loading sales order {0} line {1} - from purchase order {2}",
                    this.DropShipSalesOrderId,
                    this.DropShipSalesOrderLineId,
                    this.Document.DocNum);
                ThreadedAppLog.WriteLine("        " + task);

                return new SalesOrderAdapter(this.Company, this.DropShipSalesOrderId);
            }

            return null;
        }

        /// <summary>
        /// Queues an email to be sent at a later date
        /// </summary>
        public void QueueEmailToVendor()
        {
            string emailSubject;
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "POEmailSubject", out emailSubject);
            emailSubject = emailSubject.Replace("#DocEntry#", this.Document.DocEntry.ToString());

            string emailBody;
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "POEmailBody", out emailBody);
            emailBody = emailBody.Replace("#DocEntry#", this.Document.DocEntry.ToString());

            if (this.GetContact() != null)
            {
                this.QueueSendDocumentInEmail(emailSubject, emailBody, this.GetContact().E_Mail);
            }
        }

        /// <summary>
        /// Sends the Purchase Order to the Vendor Contact as a PDF.
        /// </summary>
        public void SendEmailToVendor()
        {
            string emailSubject;
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "POEmailSubject", out emailSubject);

            string emailBody;
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "POEmailBody", out emailBody);

            string attachment = this.CreateReport();
            this.EmailDocument(this.GetContact().E_Mail, emailSubject, emailBody, new[] { attachment });
        }

        /// <summary>
        /// Gets the payment method.
        /// </summary>
        /// <returns>The default payment method for the vendor</returns>
        public string GetPaymentMethod()
        {
            using (var rs = new RecordsetAdapter(this.Company, string.Format("SELECT PymCode FROM OCRD WHERE CardCode = '{0}'", this.Document.CardCode)))
            {
                return rs.EoF ? string.Empty : rs.FieldValue("PymCode").ToString();
            }
        }

        /// <summary>
        /// Gets the transportation code.
        /// </summary>
        /// <returns>
        /// The Transportation Code
        /// </returns>
        private int GetTransportationCode()
        {
            int transportationCode;
            ConfigurationHelper.GlobalConfiguration.Load(Company, "POTransportCode", out transportationCode);
            if (transportationCode == 0)
            {
                transportationCode = -1;
            }

            return transportationCode;
        }

        #endregion Methods
    }    
}
