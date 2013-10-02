//-----------------------------------------------------------------------
// <copyright file="SalesOrderAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------

using B1C.SAP.DI.Constants;

namespace B1C.SAP.DI.BusinessAdapters.Sales
{
    using System;
    using System.Text;
    using BusinessPartners;
    using Helpers;
    using Purchasing;
    using SAPbobsCOM;
    using Utility.Logging;
    using System.Collections.Generic;

    /// <summary>
    /// Instance of the OrderAdapter class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public class SalesOrderAdapter : ARDocumentAdapter
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="orderId">The order id.</param>
        public SalesOrderAdapter(Company company, int orderId)
            : base(company, orderId)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        public SalesOrderAdapter(Company company) : base(company)
        {
            this.Document = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="salesOrder">The sales order.</param>
        public SalesOrderAdapter(Company company, Documents salesOrder) : base(company, salesOrder)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        public SalesOrderAdapter(Company company, string xmlFile, int index)
            : base(company, xmlFile, index)
        {
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// The type of document (SalesOrder)
        /// </summary>
        public override BoObjectTypes DocumentType
        {
            get { return BoObjectTypes.oOrders; }
        }

        /// <summary>
        /// Gets the target document type of this adapter's document type
        /// </summary>
        /// <value></value>
        public override BoObjectTypes TargetType
        {
            get { return BoObjectTypes.oDeliveryNotes; }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Instances the specified parent add on.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="orderId">The order id.</param>
        /// <returns>An instance to the OrderAdapter class.</returns>
        public static SalesOrderAdapter Instance(Company company, int orderId)
        {
            return new SalesOrderAdapter(company, orderId);
        }

        /// <summary>
        /// Instances the specified parent add on.
        /// </summary>
        /// <param name="company">The SAP Company object.</param>
        /// <returns>An instance to the OrderAdapter class.</returns>
        public static SalesOrderAdapter Instance(Company company)
        {
            return new SalesOrderAdapter(company);
        }

        /// <summary>
        /// Gets all Production Order Ids associated with this Sales Order
        /// </summary>
        /// <returns>A list of DocEntrys for each Production Order referencing this Sales Order</returns>
        public IList<int> GetProductionOrderIds()
        {
            IList<int> poIds = new List<int>();

            string poNumberQuery = string.Format("SELECT DocEntry FROM OWOR where OriginNum = {0} AND Status <> 'C'", this.Document.DocNum);
            using (var poRecordset = new RecordsetAdapter(this.Company, poNumberQuery))
            {
                while (!poRecordset.EoF)
                {
                    poIds.Add((int)poRecordset.FieldValue("DocEntry"));
                    poRecordset.MoveNext();
                }
            }

            return poIds;
        }


        /// <summary>
        /// Checks if this Sales Order has related Production Orders
        /// </summary>
        /// <returns>true if the sales order has a referenced production order. False otherwise.</returns>
        public bool HasProductionOrder()
        {
            bool poFound = false;
            string poNumberQuery = string.Format("SELECT DocEntry FROM OWOR where OriginNum = {0} AND Status <> 'C'", this.Document.DocNum);
            using (var poRecordset = new RecordsetAdapter(this.Company, poNumberQuery))
            {
                if (poRecordset.RecordCount > 0)
                {
                    poFound = true;
                }
            }

            return poFound;
        }

        /// <summary>
        /// Checks if the given line is a standard Bill of Materials
        /// </summary>
        /// <param name="itemCode">The item to check</param>
        /// <returns>true if the item on the given line is a standard BOM; false otherwise</returns>
        public bool IsStandardBOM(string itemCode)
        {
            //string query = string.Format("SELECT ItemCode FROM OITM WHERE ItemCode = '{0}' AND TreeType = 'P'", itemCode);
            string query = string.Format("SELECT 1 FROM [dbo].[OITM] T1 JOIN [dbo].[OITT] T0 ON T0.[Code] = T1.[ItemCode] WHERE T1.[Canceled] <> 'Y' AND T1.ItemCode = '{0}'", itemCode);

            using (var recordset = new RecordsetAdapter(this.Company, query))
            {
                return (!recordset.EoF);
            }

        }

        /// <summary>
        /// Determines whether the given line is a standard BOM.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns>
        /// 	<c>true</c> if the line is a standard BOM; otherwise, <c>false</c>.
        /// </returns>
        public bool IsLineStandardBOM(int lineNum)
        {
            return this.IsStandardBOM(this.GetLine(lineNum).ItemCode);
        }



        /// <summary>
        /// Gets the line quantity.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns></returns>
        public decimal GetLineQuantity(int lineNum)
        {
            return (decimal)this.GetLine(lineNum).Quantity;
        }

        /// <summary>
        /// Copies to deliveryNote.
        /// </summary>
        /// <returns>The deliveryNote copied</returns>
        public DeliveryAdapter CopyToDeliveryNote()
        {
            return new DeliveryAdapter(this.Company, CopyAllToDocument(BoObjectTypes.oDeliveryNotes));
        }

        /// <summary>
        /// Copies the lines to deliverynote.
        /// </summary>
        /// <param name="lines">The lines.</param>
        /// <returns>The deliveryNote copied</returns>
        public DeliveryAdapter CopyLinesToDeliveryNote(int[] lines)
        {
            Documents deliveryNote = this.CopyLinesToDocument(BoObjectTypes.oDeliveryNotes, lines);
            return new DeliveryAdapter(this.Company, deliveryNote);
        }

        #endregion Methods
    }
}
