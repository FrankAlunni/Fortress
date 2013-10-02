//-----------------------------------------------------------------------
// <copyright file="FortressProductionOrderAdapter.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
// <web>www.b1computing.com</web>
//-----------------------------------------------------------------------

namespace Fortress.DI.BusinessAdapters.Production
{

    #region Using Directive(s)

    using B1C.SAP.DI.BusinessAdapters.Production;
    using SAPbobsCOM;
    using System.Collections.Generic;
    using B1C.SAP.DI.BusinessAdapters;
    using Fortress.DI.Model.ProductionOrder;
    using Fortress.DI.Model.Exceptions;
    using System;

    #endregion Using Directive(s)

    public class FortressProductionOrderAdapter : ProductionOrderAdapter
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="ProductionOrderAdapter"/> class, with a blank (new, unsaved) document 
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        public FortressProductionOrderAdapter(Company company) : base(company)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ProductionOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="document">The document</param>
        public FortressProductionOrderAdapter(Company company, ProductionOrders document)
            : base(company, document)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ProductionOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="docEntry">The document id</param>
        public FortressProductionOrderAdapter(Company company, int docEntry) : base(company, docEntry)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ProductionOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index of the object to load from the XML.</param>
        public FortressProductionOrderAdapter(Company company, string xmlFile, int index)
            : base(company, xmlFile, index)
        {
        }

        #region Properties

        /// <summary>
        /// Gets or sets the parent of this production order.
        /// </summary>
        /// <value>
        /// The parent.
        /// </value>
        public int Parent
        {
            get
            {
                int parent;
                if (int.TryParse(this.GetUserDefinedField("U_XX_Parent").ToString(), out parent))
                    return parent;

                return 0;
            }

            set
            {
                this.SetUserDefinedField("U_XX_Parent", value.ToString());
            }
        }

        /// <summary>
        /// Gets or sets the sales order line number.
        /// </summary>
        /// <value>
        /// The sales order line number.
        /// </value>
        public int SalesOrderLineNumber
        {
            get
            {
                int value;
                if (int.TryParse(this.GetUserDefinedField("U_XX_SOLineNUm").ToString(), out value))
                    return value;
                return 0;
            }
            set
            {
                this.SetUserDefinedField("U_XX_SOLineNUm", value.ToString());
            }
        }

        /// <summary>
        /// Gets or sets the sales order line number.
        /// </summary>
        /// <value>
        /// The sales order line number.
        /// </value>
        public int ProductionOrderLineNumber
        {
            get
            {
                int value;
                if (int.TryParse(this.GetUserDefinedField("U_XX_ParentLine").ToString(), out value))
                    return value;
                return 0;
            }
            set
            {
                this.SetUserDefinedField("U_XX_ParentLine", value.ToString());
            }
        }

        /// <summary>
        /// Gets or sets the sales order line number.
        /// </summary>
        /// <value>
        /// The sales order line number.
        /// </value>
        public int ProductionOrderNumber
        {
            get
            {
                int value;
                if (int.TryParse(this.GetUserDefinedField("U_XX_Parent").ToString(), out value))
                    return value;
                return 0;
            }
            set
            {
                this.SetUserDefinedField("U_XX_Parent", value.ToString());
            }
        }

        /// <summary>
        /// Gets or sets the revision.
        /// </summary>
        /// <value>
        /// The revision.
        /// </value>
        public int Revision
        {
            get
            {
                int value;
                if (int.TryParse(this.GetUserDefinedField("U_XX_DrawingVer").ToString(), out value))
                    return value;
                return 0;
            }
            set
            {
                this.SetUserDefinedField("U_XX_DrawingVer", value.ToString());
            }
        }

        /// <summary>
        /// Gets the line.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns>The Document Lines</returns>
        public ProductionOrders_Lines GetLine(int lineNum)
        {
            ProductionOrders_Lines docLines = this.Document.Lines;

            for (int lineIndex = 0; lineIndex < docLines.Count; lineIndex++)
            {
                docLines.SetCurrentLine(lineIndex);
                if (lineNum == docLines.LineNumber)
                {
                    break;
                }
            }

            return docLines;
        }

        /// <summary>
        /// Gets the children.
        /// </summary>
        public IList<int> Children
        {
            get
            {
                IList<int> children = new List<int>();

                string query = string.Format("SELECT DocEntry FROM OWOR WHERE U_XX_Parent = '{0}' ", this.Document.AbsoluteEntry);
                using (var recordset = new RecordsetAdapter(this.Company, query))
                {
                    while (!recordset.EoF)
                    {
                        children.Add(int.Parse(recordset.FieldValue("DocEntry").ToString()));
                        recordset.MoveNext();
                    }
                }

                return children;
            }
        }

        #endregion Properties

        /// <summary>
        /// Checks if this Sales Order has related Production Orders
        /// </summary>
        /// <returns>true if the sales order has a referenced production order. False otherwise.</returns>
        public bool HasProductionOrder(int lineNum)
        {
            string poNumberQuery = string.Format("SELECT * FROM OWOR WHERE U_XX_Parent = '{0}' AND U_XX_ParentLine = {1}", this.Document.AbsoluteEntry, lineNum);
            using (var poRecordset = new RecordsetAdapter(this.Company, poNumberQuery))
            {
                return (poRecordset.RecordCount > 0);
            }
        }


        public void CreateProductionOrder(int lineNum)
        {

            // Create All Production Orders inside a transaction
            //this.Company.StartTransaction();
            string itemCode = this.GetLine(lineNum).ItemNo;
            int productionOrderId = this.Document.AbsoluteEntry;

            // Add the Production Order
            using (var productionOrder = new FortressProductionOrderAdapter(this.Company))
            {
                productionOrder.Document.ItemNo = itemCode;
                productionOrder.Document.DueDate = DateTime.Now;
                productionOrder.Parent = this.Document.AbsoluteEntry;
                productionOrder.Revision = this.Revision;
                productionOrder.ProductionOrderLineNumber = lineNum;
                productionOrder.ProductionOrderNumber = productionOrderId;
                productionOrder.SalesOrderLineNumber = this.SalesOrderLineNumber;
                productionOrder.Document.ProductionOrderOriginEntry = this.Document.ProductionOrderOriginEntry;

                if (this.IsStandardBOM(itemCode))
                {
                    productionOrder.Document.ProductionOrderOrigin = BoProductionOrderOriginEnum.bopooSalesOrder;
                }
                else
                {
                    productionOrder.Document.ProductionOrderOrigin = BoProductionOrderOriginEnum.bopooManual;
                    productionOrder.Document.ProductionOrderType = BoProductionOrderTypeEnum.bopotSpecial;

                    ProductionOrders_Lines lines = productionOrder.Document.Lines;

                    lines.SetCurrentLine(lines.Count - 1);

                    lines.ItemNo = "AddBOMToSAP";
                    lines.BaseQuantity = 1;
                    lines.PlannedQuantity = 1;
                    lines.Add();

                }
                productionOrder.Add();
            }

            // Commit the transaction
            //this.Company.EndTransaction(BoWfTransOpt.wf_Commit);

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
        /// Checks if components have been issued
        /// </summary>
        /// <returns>true if components have been issued. false otherwise</returns>
        public bool AreComponentsIssued()
        {
            // Check if any are issued here
            string query = "SELECT SUM(IssuedQty) as issued FROM WOR1 WHERE DocEntry = " + this.Document.AbsoluteEntry;
            using (var recordset = new RecordsetAdapter(this.Company, query))
            {
                if (!recordset.EoF)
                {
                    decimal qty = decimal.Parse(recordset.FieldValue("issued").ToString());
                    if (qty > 0)
                    {
                        return true;
                    }

                    foreach (int childId in this.Children)
                    {
                        using (var childAdapter = new FortressProductionOrderAdapter(this.Company, childId))
                        {
                            if (childAdapter.AreComponentsIssued())
                            {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }
    }
}
