//-----------------------------------------------------------------------
// <copyright file="DeliveryAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.Sales
{
    #region Using Directive(s)
    
    using System;

    using Helpers;
    using SAPbobsCOM;
    using Inventory;

    #endregion Using Directive(s)

    /// <summary>
    /// Instance of the DeliveryAdapter class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public class DeliveryAdapter : ARDocumentAdapter
    {

        private PackageAdapter _package;

        public PackageAdapter Packages
        {
            get { return this._package ?? (this._package = new PackageAdapter(this.Company, this.Document.Packages)); }
        }

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="DeliveryAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="deliveryId">The delivery id.</param>
        public DeliveryAdapter(Company company, int deliveryId) : this(company)
        {
            this.Load(deliveryId);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DeliveryAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="deliveryNote">The deliverynote.</param>
        public DeliveryAdapter(Company company, Documents deliveryNote)
            : base(company, deliveryNote)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DeliveryAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        public DeliveryAdapter(Company company) : base(company)
        {
            this.Document = (Documents) company.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DeliveryAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        public DeliveryAdapter(Company company, string xmlFile, int index)
            : base(company, xmlFile, index)
        {
        }

        #endregion Constructors

        #region Properties
        /// <summary>
        /// Gets the type of document this is an adapter for
        /// </summary>
        /// <value></value>
        public override BoObjectTypes DocumentType
        {
            get { return BoObjectTypes.oDeliveryNotes; }
        }

        /// <summary>
        /// Gets the target document type of this adapter's document type
        /// </summary>
        /// <value></value>
        public override BoObjectTypes TargetType
        {
            get { return BoObjectTypes.oInvoices; }
        }
        #endregion Properties

        #region Methods

        /// <summary>
        /// Instances the specified parent add on.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="deliveryId">The delivery id.</param>
        /// <returns>An Instance to the deliveryadapter class</returns>
        public static DeliveryAdapter Instance(Company company, int deliveryId)
        {
            return new DeliveryAdapter(company, deliveryId);
        }

        /// <summary>
        /// Instances the specified parent add on.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <returns>An Instance to the deliveryadapter class</returns>
        public static DeliveryAdapter Instance(Company company)
        {
            return new DeliveryAdapter(company);
        }

        /// <summary>
        /// Loads the specified delivery id.
        /// </summary>
        /// <param name="deliveryId">The delivery id.</param>
        /// <returns>True if the method loads the class successfully.</returns>
        public bool Load(int deliveryId)
        {
            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the DeliveryAdapter.Load is not set.");
            }

            return this.Document.GetByKey(deliveryId);

        }

        /// <summary>
        /// Loads from order.
        /// </summary>
        /// <param name="orderId">The order id.</param>
        /// <param name="includeLines">if set to <c>true</c> [include lines].</param>
        /// <returns>True if the method loads the class successfully.</returns>
        public bool LoadFromOrder(int orderId, bool includeLines)
        {
            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the DeliveryAdapter.LoadFromOrder is not set.");
            }

            // This is to lcear the content of the current class
            this.Dispose();

            // Retrieve Order
            this.BaseAdapter = SalesOrderAdapter.Instance(this.Company, orderId);

            // Create Delivery
            var adapter = new DeliveryAdapter(this.Company);
            try
            {
                this.Document = adapter.Document;

                this.Document.CardCode = this.BaseAdapter.Document.CardCode;
                this.Document.SalesPersonCode = this.BaseAdapter.Document.SalesPersonCode;
                this.Document.DocType = BoDocumentTypes.dDocument_Items;
                this.Document.DocumentSubType = BoDocumentSubType.bod_None;
                this.Document.Address = this.BaseAdapter.Document.Address;
                this.Document.ShipToCode = this.BaseAdapter.Document.ShipToCode;
                this.Document.Address2 = this.BaseAdapter.Document.Address2;
                this.Document.HandWritten = BoYesNoEnum.tNO;
                this.Document.JournalMemo = string.Format("Delivery - {0}", this.BaseAdapter.Document.CardCode);
                this.SetComments(string.Format("Based on Sales Order {0}", orderId));
                this.Document.PaymentGroupCode = this.BaseAdapter.Document.PaymentGroupCode;
                this.Document.DocObjectCode = BoObjectTypes.oDeliveryNotes;
                this.Document.DocObjectCodeEx = BoObjectTypes.oDeliveryNotes.ToString();
                this.Document.DocDueDate = DateTime.Today;

                // Create delivery lines
                return !includeLines || this.LoadLinesFromOrder();
            }
            finally
            {
                adapter.Dispose();
            }
        }

        /// <summary>
        /// Loads the lines from order.
        /// </summary>
        /// <returns>True if the method loads the lines successfully.</returns>
        public bool LoadLinesFromOrder()
        {
            int deliveryLineNum = 0;
            var baseLines = this.BaseAdapter.Lines;

            foreach (Document_Lines line in baseLines)
            {
                this.Lines.Add(deliveryLineNum);

                this.LoadLineFromOrderLine(deliveryLineNum, line, line.Quantity);
                deliveryLineNum++;
            }

            return true;
        }

        /// <summary>
        /// Loads the line from order lines. Allow partial delivery.
        /// </summary>
        /// <param name="deliveryLineNum">The delivery line num.</param>
        /// <param name="orderLine">The order line.</param>
        /// <param name="quantity">The quantity.</param>
        /// <returns>True if the method loads the line successfully.</returns>
        public bool LoadLineFromOrderLine(int deliveryLineNum, Document_Lines orderLine, double quantity)
        {
            var deliveryLine = Lines.Current;
            try
            {
                deliveryLine.ItemCode = orderLine.ItemCode;
                deliveryLine.Quantity = quantity;
                deliveryLine.Rate = orderLine.Rate;
                deliveryLine.BaseType = (int)BoAPARDocumentTypes.bodt_Order;
                deliveryLine.BaseEntry = this.BaseAdapter.Document.DocNum;
                deliveryLine.BaseLine = orderLine.LineNum;
                return true;
            }
            finally
            {
                COMHelper.Release(ref deliveryLine);
            }
        }

        /// <summary>
        /// Creates an A/R Invoice from this DeliveryNote
        /// </summary>
        /// <remarks>
        /// The returned document is not saved to the DB. It is the responsibility of the calling
        /// object to call the Add() method on the object if it is wished to be saved.
        /// </remarks>
        /// <returns>The newly created (unsaved) A/R Invoice</returns>
        public ARInvoiceAdapter CopyToARInvoice()
        {
            var receivableInvoice = CopyAllToDocument(BoObjectTypes.oInvoices);

            var invoiceAdapter = new ARInvoiceAdapter(this.Company, receivableInvoice);
            return invoiceAdapter;
        }

        /// <summary>
        /// Systems the transportation code.
        /// </summary>
        /// <returns>Return system transportation code</returns>
        public int SystemTransportationCode()
        {
            int transportationCode;
            ConfigurationHelper.GlobalConfiguration.Load(Company, "POTransportCode", out transportationCode);
            if (transportationCode == 0)
            {
                transportationCode = -1;
            }

            return transportationCode;
        }

        protected override void Release()
        {
            if (this._package != null)
            {
                this.Packages.Dispose();
            }

            base.Release();
        }
        #endregion Methods
    }
}
