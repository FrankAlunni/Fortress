//-----------------------------------------------------------------------
// <copyright file="FortressSalesOrderAdapter.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
//-----------------------------------------------------------------------

namespace Fortress.DI.BusinessAdapters.Sales
{
    #region Using Directive(s)

    using System;
    using System.Collections.Generic;
    using System.Linq;
    using B1C.SAP.DI.BusinessAdapters;
    using B1C.SAP.DI.BusinessAdapters.Sales;
    using SAPbobsCOM;
    using System.IO;
    using Helpers;
    using Model.Exceptions;
    using Production;
    using B1C.SAP.DI.BusinessAdapters.Inventory;
    using Model.ProductionOrder;
    using Model.Enums;
    using System.Xml;
    using System.Web;
    using B1C.SAP.DI.Business;
    using B1C.SAP.DI.Helpers;

    #endregion Using Directive(s)

    public class FortressSalesOrderAdapter : SalesOrderAdapter
    {
        #region Private Member(s)

        #endregion Private Member(s)

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="FortressSalesOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="orderId">The order id.</param>
        public FortressSalesOrderAdapter(Company company, int orderId)
            : base(company, orderId)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FortressSalesOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        public FortressSalesOrderAdapter(Company company) : base(company)
        {
            this.Document = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FortressSalesOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="salesOrder">The sales order.</param>
        public FortressSalesOrderAdapter(Company company, Documents salesOrder) : base(company, salesOrder)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FortressSalesOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        public FortressSalesOrderAdapter(Company company, string xmlFile, int index)
            : base(company, xmlFile, index)
        {
        }

        #endregion Constructor(s)

        #region Properties

        #endregion Properties

        #region Method(s)

        #region Public Method(s)

        /// <summary>
        /// Checks if XML exists for the given line number/itemCode.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">The item code.</param>
        /// <returns>true if an XML file exists, false otherwise</returns>
        public bool CheckIfXmlExists(int lineNum, string itemCode)
        {
            bool exists = false;

            // Check in the Drawing Directory
            var soDir = this.GetSalesOrderLineDirectory(lineNum, itemCode);

            if (soDir != null)
            {
                // Attempt to find an XML
                string fileName = this.GetSalesOrderLineXmlFileName(lineNum, itemCode);

                FileInfo[] matchingFiles = soDir.GetFiles(fileName);

                if (matchingFiles.Length > 0)
                    exists = true;
            }

            return exists;
        }


        /// <summary>
        /// Checks if XML exists for the given line number/itemcode.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">The item code.</param>
        /// <param name="revision">The revision.</param>
        /// <returns>true if an XML file is found, false otherwise</returns>
        public bool CheckIfXmlExists(int lineNum, string itemCode, int revision)
        {
            bool exists = false;

            // Check in the Drawing Directory
            var soDir = this.GetSalesOrderLineDirectory(lineNum, itemCode);

            if (soDir != null)
            {
                // Attempt to find an XML
                string fileName = this.GetSalesOrderLineXmlFileName(lineNum, itemCode, revision);

                FileInfo[] matchingFiles = soDir.GetFiles(fileName);

                if (matchingFiles.Length > 0)
                    exists = true;
            }

            return exists;
        }

        /// <summary>
        /// Creates the production orders for one line on the sales order.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        //[Obsolete("CreateProductionOrders is deprecated, please use CreateProductionOrder instead.")]
        //public void CreateProductionOrders(int lineNum)
        //{            
        //    bool statusPassed = this.HasConfigurationPassed(lineNum);
        //    if (!statusPassed)
        //    {
        //        throw new FortressException("Configuration status is not set to Passed");
        //    }

        //    // Check the UDO statuses
        //    IList<string> statuses = this.GetLineApprovalStatuses(lineNum);
        //    ApprovalStatus lineStatus = this.GetLineApprovalStatus(lineNum);
            
        //    if (!(statusPassed && (lineStatus == ApprovalStatus.None || lineStatus == ApprovalStatus.No)))
        //    {
        //        if (!(statusPassed && (lineStatus != ApprovalStatus.None && lineStatus != ApprovalStatus.No) && (statuses.Contains("Approved") && statuses.Contains("Prod. Drawing Completed"))))
        //        {
        //            throw new FortressException("Cannot create production order. Line must first be approved and production drawings completed.");
        //        }
        //    }

        //    // Find all production orders to create
        //    RequestedProductionOrder posToCreate = this.DoPrepareProductionOrders(lineNum, this.GetLine(lineNum).ItemCode);

        //    try
        //    {
        //        IList<IDictionary<int, IList<double>>> pqsToUpdate = new List<IDictionary<int, IList<double>>>();

        //        // Create All Production Orders inside a transaction
        //        this.Company.StartTransaction();

        //        // Cancel any existing production orders
        //        IList<int> existingPos = this.GetProductionOrdersForLine(lineNum);

        //        foreach (int existingPo in existingPos)
        //        {
        //            using (var fpoa = new FortressProductionOrderAdapter(this.Company, existingPo))
        //            {
        //                fpoa.Cancel();
        //            }
        //        }
                
                
        //        // Add 1 for each quantity
        //        for (int i = 0; i < this.GetLineQuantity(lineNum); i++)
        //        {
        //            // Recursively add this production order, and all children
        //            var pqs = this.AddRequestedProductionOrder(posToCreate, null);
        //            pqsToUpdate.Add(pqs);
        //        }

        //        // Update the line status
        //        this.SetProductionOrderLineStatus(lineNum,
        //                                          posToCreate.GetType() == typeof (RequestedSpecialProductionOrder)
        //                                              ? ProductionOrderStatus.SpecialProductionOrderSuccess
        //                                              : ProductionOrderStatus.StandardProductionOrderSuccess);

        //        // Update the SO with the new status
        //        this.Update();

        //        // Commit the transaction
        //        this.Company.EndTransaction(BoWfTransOpt.wf_Commit);

        //        // Update all the planned qtys
        //        foreach (var pqDic in pqsToUpdate)
        //        {
        //            foreach (var poId in pqDic.Keys)
        //            {
        //                using (var po = new FortressProductionOrderAdapter(this.Company, poId))
        //                {
        //                    int index = 0;
        //                    IList<double> vals = pqDic[poId];
        //                    foreach (ProductionOrders_Lines line in po.Lines)
        //                    {
        //                        line.PlannedQuantity = vals[index];
        //                        index++;
        //                    }

        //                    po.Update();
        //                }
        //            }
        //        }                
        //    }
        //    catch (Exception ex)
        //    {
        //        try
        //        {
        //            // If an exception occurs, rollback transaction and rethrow
        //            this.Company.EndTransaction(BoWfTransOpt.wf_RollBack);
        //        }
        //        catch (Exception) { }

        //        // Update the line status
        //        this.SetProductionOrderLineStatus(lineNum,
        //                                          posToCreate.GetType() == typeof (RequestedSpecialProductionOrder)
        //                                              ? ProductionOrderStatus.SpecialProductionOrderFailed
        //                                              : ProductionOrderStatus.StandardProductionOrderFailed);

        //        // Update the SO with the new status
        //        this.Update();

        //        throw ex;
        //    }
        //}

        public void CreateProductionOrder(int lineNum)
        {
            // Validate the configuration 
            string itemCode = string.Empty;
            string query = string.Format("SELECT ItemCode FROM RDR1 WHERE DocEntry = {0} AND LineNum = {1}", this.DocEntry, lineNum);
            using (var rec = new RecordsetAdapter(this.Company, query))
            {
                if (!rec.EoF)
                {
                    itemCode = rec.FieldValue("ItemCode").ToString();
                }
            }

            if ((this.IsMetalDetector(itemCode)) || (this.IsConveyorBelt(itemCode)))
            {
                if (!this.HasConfigurationPassed(lineNum))
                {
                    throw new FortressException("Configuration status is not set to Passed");
                }
            }


            // Check the UDO statuses
            IList<string> statuses = this.GetLineApprovalStatuses(lineNum);
            ApprovalStatus lineStatus = this.GetLineApprovalStatus(lineNum);

            if (!(lineStatus != ApprovalStatus.None && lineStatus != ApprovalStatus.No) && (statuses.Contains("Approved") && statuses.Contains("Prod. Drawing Completed")))
            {
                throw new FortressException("Cannot create production order. Line must first be approved and production drawings completed.");
            }

            // Find the production order to create
            RequestedProductionOrder request = this.PrepareRequest(lineNum, this.GetLine(lineNum).ItemCode, this.GetLine(lineNum).Quantity);

            try
            {
                // Create All Production Orders inside a transaction
                this.Company.StartTransaction();

                // Add the Production Order
                if (request.GetType() == typeof(RequestedStandardProductionOrder))
                {
                    this.AddStandardProductionOrder((RequestedStandardProductionOrder)request);
                }
                else if (request.GetType() == typeof(RequestedAddBomProductionOrder))
                {
                    this.AddAddBomProductionOrder((RequestedAddBomProductionOrder)request);
                }
                else
                {
                    throw new FortressException("Unknown request.");
                }

                // Update the line status
                this.SetProductionOrderLineStatus(lineNum,
                                                  request.GetType() == typeof(RequestedSpecialProductionOrder)
                                                      ? ProductionOrderStatus.SpecialProductionOrderSuccess
                                                      : ProductionOrderStatus.StandardProductionOrderSuccess);

                // Update the SO with the new status
                this.Update();

                // Commit the transaction
                this.Company.EndTransaction(BoWfTransOpt.wf_Commit);


            }
            catch (Exception ex)
            {
                try
                {
                    // If an exception occurs, rollback transaction and rethrow
                    this.Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                catch (Exception) { }

                // Update the line status
                this.SetProductionOrderLineStatus(lineNum,
                                                  request.GetType() == typeof(RequestedSpecialProductionOrder)
                                                      ? ProductionOrderStatus.SpecialProductionOrderFailed
                                                      : ProductionOrderStatus.StandardProductionOrderFailed);

                // Update the SO with the new status
                this.Update();

                throw ex;
            }
        }


        /// <summary>
        /// Gets the current revision of the given line.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns>The current revision number</returns>
        public int GetCurrentLineRevision(int lineNum)
        {            
            int revision = -1;

            var query = string.Format(@"select ISNULL(MAX(CAST(U_XX_Revis as INT)), -1) as revision FROM
                                    (
                                    select ISNULL(U_XX_Revis, '0') as U_XX_Revis FROM [@XX_CONVBLT] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}
                                    union all
                                    select ISNULL(U_XX_Revis, '0') as U_XX_Revis FROM [@XX_METDET] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}
                                    ) as revisions", this.Document.DocEntry, lineNum);

            using (var revRecordset = new RecordsetAdapter(this.Company, query))
            {
                if (!revRecordset.EoF)
                {
                    revision = int.Parse(revRecordset.FieldValue("revision").ToString());
                }
            }

            return revision;
        }


        /// <summary>
        /// Determines whether the current line has been configured (Conveyor Belt or Metal Detector).
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns>
        /// 	<c>true</c> if the line has been configured; otherwise, <c>false</c>.
        /// </returns>
        public bool IsLineConfigured(int lineNum)
        {
            string itemCode = string.Empty;
            string query = string.Format("SELECT ItemCode FROM RDR1 WHERE DocEntry = {0} AND LineNum = {1}", this.DocEntry, lineNum);
            using (var rec = new RecordsetAdapter(this.Company, query))
            {
                if (!rec.EoF)
                {
                    itemCode = rec.FieldValue("ItemCode").ToString();
                }
            }

            if ((this.IsMetalDetector(itemCode)) || (this.IsConveyorBelt(itemCode)))
            {
                // It will have a revision of -1 if there is no configuration
                return this.GetCurrentLineRevision(lineNum) != -1;
            }

            return true;
        }

        /// <summary>
        /// Gets the production order for line.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns></returns>
        public IList<int> GetProductionOrdersForLine(int lineNum)
        {
            IList<int> pos = new List<int>();

            string poNumberQuery = string.Format("SELECT DocEntry FROM OWOR where OriginNum = {0} AND Status <> 'C' AND U_XX_SOLineNUm = {1} order by U_XX_DrawingVer desc, DocEntry asc", this.Document.DocNum, lineNum);
            using (var poRecordset = new RecordsetAdapter(this.Company, poNumberQuery))
            {
                while (!poRecordset.EoF)
                {
                    pos.Add((int)poRecordset.FieldValue("DocEntry"));
                    poRecordset.MoveNext();
                }
            }

            return pos;
        }

        /// <summary>
        /// Checks if this Sales Order Line has related Production Orders
        /// </summary>
        /// <returns>true if the sales order has a referenced production order. False otherwise.</returns>
        public bool HasProductionOrder(int lineNum)
        {
            bool poFound = false;
            string poNumberQuery = string.Format("SELECT DocEntry FROM OWOR where OriginNum = {0} AND Status <> 'C' AND U_XX_SOLineNUm = {1}", this.Document.DocNum, lineNum);
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
        /// Gets the name of the sales order line XML file.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">the item code</param>
        /// <returns>the name of the XML file name for the given linenum/itemcode</returns>
        public string GetSalesOrderLineXmlFileName(int lineNum, string itemCode)
        {
            string dirPattern = FortressConfigurationHelper.GetBomFilePattern(this.Company);

            dirPattern = ApplyPattern(dirPattern);
            dirPattern = ApplyPattern(lineNum, itemCode, dirPattern);

            return dirPattern;
        }

        /// <summary>
        /// Gets the name of the sales order line XML file.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">the item code</param>
        /// <param name="revision">The revision.</param>
        /// <returns>the name of the XML file for the given linenum/itemcode/revision</returns>
        public string GetSalesOrderLineXmlFileName(int lineNum, string itemCode, int revision)
        {
            string dirPattern = FortressConfigurationHelper.GetBomFilePattern(this.Company);

            dirPattern = ApplyPattern(dirPattern);
            dirPattern = ApplyPattern(lineNum, itemCode, revision, dirPattern);

            return dirPattern;
        }

        /// <summary>
        /// Gets the sales order line directory.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">the item code</param>
        /// <returns>The directory for the given sales order</returns>
        public DirectoryInfo GetSalesOrderLineDirectory(int lineNum, string itemCode)
        {
            string dirPattern = FortressConfigurationHelper.GetLineDirectoryPattern(this.Company);

            dirPattern = ApplyPattern(dirPattern);
            dirPattern = ApplyPattern(lineNum, itemCode, dirPattern);

            DirectoryInfo dInfo = this.GetSalesOrderDirectory(lineNum);
            DirectoryInfo lineDir = new DirectoryInfo(dInfo.FullName + '\\' + dirPattern);

            if (!lineDir.Exists)
            {
                return null;
            }

            return lineDir;
        }


        /// <summary>
        /// Gets or creates (if it doesn't exist the sales order line directory.
        /// DOM - 110326
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">the item code</param>
        /// <returns>The directory for the given sales order</returns>
        public DirectoryInfo GetOrCreateSalesOrderLineDirectory(int lineNum, string itemCode)
        {
            
            DirectoryInfo lineDir = GetSalesOrderLineDirectory(lineNum,itemCode);

            if (lineDir != null)
            {
                CreateSalesOrderLineDirectory(lineNum, itemCode); // Create it
                lineDir = GetSalesOrderLineDirectory(lineNum, itemCode);  // Get it
            }

            return lineDir;
        }





        /// <summary>
        /// Creates the sales order line directory.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">The item code.</param>
        /// <returns></returns>
        public DirectoryInfo CreateSalesOrderLineDirectory(int lineNum, string itemCode)
        {
            string dirPattern = FortressConfigurationHelper.GetLineDirectoryPattern(this.Company);

            dirPattern = ApplyPattern(dirPattern);
            dirPattern = ApplyPattern(lineNum, itemCode, dirPattern);

            DirectoryInfo dInfo = this.GetSalesOrderDirectory(lineNum);
            DirectoryInfo lineDir = new DirectoryInfo(dInfo.FullName + '\\' + dirPattern);
            if (!lineDir.Exists)
            {
                lineDir.Create();
                lineDir = new DirectoryInfo(this.GetSalesOrderDirectory(lineNum).FullName + "/" + dirPattern);
            }

            return lineDir;
        }

        public void CreateAllSalesOrderLineDirectories()
        {
            foreach (Document_Lines line in this.Lines)
            {
                this.CreateSalesOrderLineDirectory(line.LineNum, line.ItemCode);
            }
        }

        /// <summary>
        /// Gets the sales order directory.
        /// </summary>
        /// <returns></returns>
        public DirectoryInfo GetSalesOrderDirectory(int lineNum)
        {
            /*
            string dirPattern = FortressConfigurationHelper.GetSalesOrderDirectoryPattern(this.Company);

            dirPattern = ApplyPattern(dirPattern);

            this.SetLine(lineNum);
            string soDir = FortressConfigurationHelper.GetDrawingDirectory(this.Company).FullName;
            */

            return GetOrCreateSalesOrderDirectory(this.Company, this.DocEntry);
            //return GetOrCreateSalesOrderDirectory(this.Company, this.DocNumber);
        }

        /// <summary>
        /// Duplicates the configuration.
        /// </summary>
        /// <param name="fromLine">From line.</param>
        /// <param name="toLine">To line.</param>
        public void DuplicateConfiguration(int fromLine, int toLine)
        {
            this.DuplicateConfiguration(fromLine, this.DocEntry, toLine);
        }

        /// <summary>
        /// Duplicates the configuration from the fromLine to the toLine.
        /// </summary>
        /// <param name="fromLine">From line.</param>
        /// <param name="toDocEntry">To doc entry.</param>
        /// <param name="toLine">To line.</param>
        public void DuplicateConfiguration(int fromLine, int toDocEntry, int toLine)
        {
            string docEntryQuery = string.Format("SELECT DocEntry FROM ORDR Where DocEntry = {0}", toDocEntry);
            using (var rec = new RecordsetAdapter(this.Company, docEntryQuery))
            {
                if (rec.EoF)
                {
                    throw new FortressException(string.Format("No Sales order with DocEntry {0} can be found. Cant duplicate configuration.", toDocEntry));
                }
            }

            string itemCode = string.Empty;
            string query = string.Format("SELECT ItemCode FROM RDR1 WHERE DocEntry = {0} AND LineNum = {1}", this.DocEntry, fromLine);
            using (var rec = new RecordsetAdapter(this.Company, query))
            {
                if (!rec.EoF)
                {
                    itemCode = rec.FieldValue("ItemCode").ToString();
                }
            }

            if (this.IsMetalDetector(itemCode))
            {
                this.DuplicateMetalDetectorConfiguration(fromLine, toDocEntry,  toLine);
            }
            else if (this.IsConveyorBelt(itemCode))
            {
                this.DuplicateConveyorBeltConfiguration(fromLine, toDocEntry,  toLine);
            }

            // Make sure all directories are present
            this.CreateAllSalesOrderLineDirectories();
        }



        #endregion Public Method(s)

        #region Private Method(s)

        private bool IsConveyorBelt(string itemCode)
        {
            string prefix;
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "ConvPref", out prefix);

            if (!itemCode.StartsWith(prefix))
            {
                return false;
            }
            return true;
        }

        private bool IsMetalDetector(string itemCode)
        {
            string prefix;
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "DetPref", out prefix);

            if (!itemCode.StartsWith(prefix))
            {
                return false;
            }
            return true;
        }

        private void DuplicateMetalDetectorConfiguration(int fromLine, int toDocEntry, int toLine)
        {
            // Duplicate row
            using (var rs = new RecordsetAdapter(this.Company, string.Format("SELECT * FROM [@XX_METDET] WHERE U_XX_OrderNo = '{0}' and U_XX_OrdrLnNo = '{1}'", this.DocEntry, fromLine)))
            {
                if (!rs.EoF)
                {
                    
                    string newCode = toDocEntry.ToString() + '-' + toLine.ToString();

                    var loadData = new System.Text.StringBuilder();
                    loadData.AppendLine("<DataLoader UDOName='MDConfigurator' UDOType='MasterData'>");
                    loadData.AppendLine("<Record>");
                    loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(newCode));
                    loadData.AppendFormat("<Field Name='U_XX_OrderNo' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(toDocEntry.ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OrdrLnNo' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(toLine.ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Revis' Type='System.String' Value='{0}' />", "1");
                    loadData.AppendFormat("<Field Name='U_XX_Status' Type='System.String' Value='{0}' />", "Initiated");
                    loadData.AppendFormat("<Field Name='U_XX_Descr' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Descr").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_PkgMat' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_PkgMat").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_PrdPerMn' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_PrdPerMn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Applic' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Applic").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_PkgSh' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_PkgSh").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BlkFlRt' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BlkFlRt").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_TempU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_TempU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Temp' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Temp").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdWdMn' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdWdMn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdWdMx' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdWdMx").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdHgMn' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdHgMn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdHgMx' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdHgMx").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdLnMn' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdLnMn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdLnMx' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdLnMx").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdDmMn' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdDmMn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdDmMx' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdDmMx").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdWdMnU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdWdMnU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdWdMxU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdWdMxU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdHgMnU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdHgMnU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdHgMxU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdHgMxU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdLnMnU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdLnMnU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdLnMxU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdLnMxU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdDmMnU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdDmMnU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdDmMxU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdDmMxU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdWgUn' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdWgUn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdWgMx' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdWgMx").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ProdWgMn' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ProdWgMn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_WebWd' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_WebWd").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_TrackEr' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_TrackEr").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Thick' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Thick").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_FeedRt' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_FeedRt").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Densit' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Densit").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_TabletWd' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_TabletWd").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_TabletWdU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_TabletWdU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_TabletHg' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_TabletHg").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_TabletHgU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_TabletHgU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_TabletSh' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_TabletSh").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_IronEnr' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_IronEnr").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Viscos' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Viscos").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Press' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Press").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Model' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Model").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_PartNo' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_PartNo").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BeltWd' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BeltWd").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BeltWdU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BeltWdU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_MDUnit' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_MDUnit").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BSH' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BSH").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_CtrlTyp' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_CtrlTyp").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Style' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Style").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ApertWd' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ApertWd").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ApertWdU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ApertWdU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ApertHg' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ApertHg").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ApertHgU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ApertHgU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ApertDm' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ApertDm").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ApertDmU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ApertDmU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ApertTC' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ApertTC").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ApertBC' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ApertBC").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Freq1' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Freq1").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_DualFreq' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_DualFreq").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Freq2' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Freq2").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_DetType' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_DetType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_DetFnsh' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_DetFnsh").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_PhMode' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_PhMode").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Enviro' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Enviro").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Lang' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Lang").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_DetFlow' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_DetFlow").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_DetChute' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_DetChute").ToString()));

                    loadData.AppendFormat("<Field Name='U_XX_SensitFE' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_SensitFE").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_SensitNF' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_SensitNF").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_SensitSS' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_SensitSS").ToString()));

                    loadData.AppendFormat("<Field Name='U_XX_SensDir' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_SensDir").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Shield' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Shield").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_MultiSc' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_MultiSc").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_AlarmBR' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_AlarmBR").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_AlarmMBO' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_AlarmMBO").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_AlarmRR' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_AlarmRR").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RgRequ' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RgRequ").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RgSize' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RgSize").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RgSizeU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RgSizeU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RgContct' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RgContct").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RgContTp' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RgContTp").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_ExplPrf' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_ExplPrf").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Remote' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Remote").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RemoteTp' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RemoteTp").ToString()));

                    loadData.AppendFormat("<Field Name='U_XX_UNotes' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_UNotes").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_SNotes' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_SNotes").ToString()));

                    loadData.AppendLine("</Record>");
                    loadData.AppendLine("</DataLoader>");

                    DataLoader loader = DataLoader.InstanceFromXMLString(loadData.ToString());
                    if (loader != null)
                    {
                        loader.Process(this.Company);

                        if (!string.IsNullOrEmpty(loader.ErrorMessage))
                        {
                            throw new FortressException("Data loader exception: " + loader.ErrorMessage);
                        }
                        else
                        {
                            this.DuplicateCommission(fromLine, toDocEntry, toLine);
                        }
                    }
                }
                else
                {
                    this.DuplicateCommission(fromLine, toDocEntry, toLine);
                }
            }
        }

        private void DuplicateConveyorBeltConfiguration(int fromLine, int toDocEntry, int toLine)
        {
            // Duplicate row
            using (var rs = new RecordsetAdapter(this.Company, string.Format("SELECT * FROM [@XX_CONVBLT] WHERE U_XX_OrderNo = '{0}' and U_XX_OrdrLnNo = '{1}'", this.DocEntry, fromLine)))
            {
                if (!rs.EoF)
                {

                    string newCode = toDocEntry.ToString() + '-' + toLine.ToString();

                    var loadData = new System.Text.StringBuilder();
                    loadData.AppendLine("<DataLoader UDOName='CBConfigurator' UDOType='MasterData'>");
                    loadData.AppendLine("<Record>");
                    loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(newCode));
                    loadData.AppendFormat("<Field Name='U_XX_OrderNo' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(toDocEntry.ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OrdrLnNo' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(toLine.ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Revis' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode("1"));
                    loadData.AppendFormat("<Field Name='U_XX_Status' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode("Initiated"));
                    loadData.AppendFormat("<Field Name='U_XX_OvLenght' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OvLenght").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltType' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltSpeed' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltSpeed").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltWd' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltWd").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltPM' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltPM").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltLn' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltLn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltDir' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltDir").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltHgIn' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltHgIn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltHgOut' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltHgOut").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltAdj' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltAdj").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_DetLoc' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_DetLoc").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_DetLocC' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_DetLocC").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_DetFrCtr' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_DetFrCtr").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_MntType' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_MntType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_MntStyle' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_MntStyle").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_MntMat' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_MntMat").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutRReq' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutRReq").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutRStl' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutRStl").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutLn' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutLn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutLnU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutLnU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutWd' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutWd").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutWdU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutWdU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutHg' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutHg").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutHgU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutHgU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_FrmType' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_FrmType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_CtrlLoc' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_CtrlLoc").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_MtrReq' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_MtrReq").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_SngDrop' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_SngDrop").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Voltage' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Voltage").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Phase' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Phase").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_HP' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_HP").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_MotType' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_MotType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_HrzMot' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_HrzMot").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_SlvDrive' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_SlvDrive").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_GearBD' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_GearBD").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RejReq' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RejReq").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Unit' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Unit").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_EmergStop' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_EmergStop").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_InFeedEy' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_InFeedEy").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_EncReq' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_EncReq").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_TraBead' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_TraBead").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_TraSide' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_TraSide").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BmpGrd' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BmpGrd").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Alarm' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_Alarm").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RailType' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RailType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RailLn' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RailLn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RailHg' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RailHg").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RailHgOpt' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RailHgOpt").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_StType' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_StType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_StOptPt' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_StOptPt").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_StOptKR' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_StOptKR").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_StOptBR' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_StOptBR").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_StOptBS' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_StOptBS").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_StOptES' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_StOptES").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_LatcType' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_LatcType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_LSType' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_LSType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_LSLoc' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_LSLoc").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_LSRed' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_LSRed").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_LSAmber' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_LSAmber").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_LSGreen' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_LSGreen").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_LSLtc' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_LSLtc").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_LSBuzzer' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_LSBuzzer").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RejTyp1' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RejTyp1").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RejRet' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RejRet").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_SecRej' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_SecRej").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RejTyp2' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RejTyp2").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RejDir' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RejDir").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RejAlrm' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RejAlrm").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RejBinTp' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RejBinTp").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_RejBinSt' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_RejBinSt").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BSAPot' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BSAPot").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BSAKeyR' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BSAKeyR").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BSABltR' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BSABltR").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BSABtSw' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BSABtSw").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BSAEmSt' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BSAEmSt").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_LatcType2' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_LatcType2").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_AirBTR' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_AirBTR").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_CapTank' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_CapTank").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_AirPrS' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_AirPrS").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_KickHD' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_KickHD").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_KickFM' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_KickFM").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_FlGtLn' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_FlGtLn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_FlChute' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_FlChute").ToString()));

                    loadData.AppendFormat("<Field Name='U_XX_UNotes' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_UNotes").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_SNotes' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_SNotes").ToString()));

                    loadData.AppendLine("</Record>");
                    loadData.AppendLine("</DataLoader>");

                    DataLoader loader = DataLoader.InstanceFromXMLString(loadData.ToString());
                    if (loader != null)
                    {
                        loader.Process(this.Company);

                        if (!string.IsNullOrEmpty(loader.ErrorMessage))
                        {
                            throw new Exception("Data loader exception: " + loader.ErrorMessage);
                        }
                        else
                        {
                            this.DuplicateCommission(fromLine, toDocEntry, toLine);
                        }
                    }
                }
                else
                {
                    this.DuplicateCommission(fromLine, toDocEntry, toLine);
                }
            }

        }

        public void DuplicateCommission(int fromLine, int toLine)
        {
            this.DuplicateCommission(fromLine, this.DocEntry, toLine);
        }

        public void DuplicateCommission(int fromLine, int toDocEntry, int toLine)
        {

            var loadData = new System.Text.StringBuilder();
            loadData.AppendLine("<DataLoader UDOName='Commission' UDOType='MasterData'>");


            // Get all the data from the current commision:
            string query = string.Format(@"SELECT DocEntry, Code, U_XX_Type as ComType, Name, U_XX_Rate as Rate, U_XX_CommisAv as commis, CONVERT(varchar(10), CreateDate, 121) as CreateDate, LEFT(CreateTime, len(CreateTime)-2) + ':' + RIGHT(CreateTime, 2) as CreateTime, CONVERT(varchar(10), UpdateDate, 121) as UpdateDate, LEFT(UpdateTime, len(UpdateTime)-2) + ':' + RIGHT(UpdateTime, 2) as UpdateTime, U_XX_Approved  as Approved FROM [@XX_COMMISS] WITH (NOLOCK) WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}", this.DocEntry, fromLine);

            using (var recordset = new RecordsetAdapter(this.Company, query))
            {
                // If there are no commission lines, return.
                if (recordset.EoF)
                {
                    return;
                }

                int counter = 0;
                while (!recordset.EoF)
                {
                    string nextCode = new TableFieldAdapter(this.Company, "@XX_COMMISS", "Code").NextCode("CMM", string.Empty, counter);

                    loadData.AppendLine("<Record>");
                    loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(nextCode));
                    loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(recordset.FieldValue("Name").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Rate' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(recordset.FieldValue("Rate").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Type' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(recordset.FieldValue("ComType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OrderNo' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(toDocEntry.ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OrdrLnNo' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(toLine.ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_Approved' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(recordset.FieldValue("Approved").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_CommisAv' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(recordset.FieldValue("commis").ToString()));
                    loadData.AppendLine("</Record>");

                    counter++;
                    recordset.MoveNext();
                }
            }

            loadData.AppendLine("</DataLoader>");

            DataLoader loader = DataLoader.InstanceFromXMLString(loadData.ToString());
            if (loader != null)
            {
                loader.Process(this.Company);

                if (!string.IsNullOrEmpty(loader.ErrorMessage))
                {
                    throw new FortressException("Data loader exception: " + loader.ErrorMessage);
                }
            }


            return;
        }


        /// <summary>
        /// Adds the requested production order.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <param name="parent">The parent.</param>
        private IDictionary<int, IList<double>> AddRequestedProductionOrder(RequestedProductionOrder request, FortressProductionOrderAdapter parent)
        {
            IDictionary<int, IList<double>> plannedQtyUpdates = new Dictionary<int, IList<double>>();

            if (request.GetType() == typeof(RequestedSpecialProductionOrder))
            {
                int poNum;
                var pqs = this.AddSpecialProductionOrder((RequestedSpecialProductionOrder)request, parent, out poNum);

                plannedQtyUpdates[poNum] = pqs;
            }
            else if (request.GetType() == typeof(RequestedStandardProductionOrder))
            {
                this.AddStandardProductionOrder((RequestedStandardProductionOrder)request, parent);
            }
            else
            {
                throw new FortressException(string.Format("Type {0} is not a supported RequestedProductionOrder", request.GetType()));
            }

            return plannedQtyUpdates;
        }

        /// <summary>
        /// Adds the special production order.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <param name="parent">The parent.</param>
        private IList<double> AddSpecialProductionOrder(RequestedSpecialProductionOrder request, FortressProductionOrderAdapter parent, out int prodOrderId)
        {
            IList<double> pqUpdates = new List<double>();

            // Attempt to load the Special production order from file, and add
            using (var productionOrder = new FortressProductionOrderAdapter(this.Company, request.FileLocation, request.XmlIndex))
            {
                if (parent != null)
                {
                    productionOrder.Parent = parent.Document.AbsoluteEntry;
                    productionOrder.Revision = parent.Revision;
                    productionOrder.SalesOrderLineNumber = parent.SalesOrderLineNumber;
                    productionOrder.Document.ProductionOrderOrigin = BoProductionOrderOriginEnum.bopooSalesOrder;
                    productionOrder.Document.ProductionOrderOriginEntry = this.Document.DocEntry;
                    productionOrder.Revision = request.DrawingVersion;
                }
                else
                {
                    productionOrder.Document.ProductionOrderOrigin = BoProductionOrderOriginEnum.bopooSalesOrder;
                    productionOrder.Document.ProductionOrderOriginEntry = this.Document.DocEntry;
                    productionOrder.SalesOrderLineNumber = request.SalesOrderLine;
                    productionOrder.Revision = request.DrawingVersion;
                }

                foreach (ProductionOrders_Lines line in productionOrder.Lines)
                {
                    pqUpdates.Add(line.PlannedQuantity);
                }

                productionOrder.Add();

                if (productionOrder.Exists)
                {
                    prodOrderId = productionOrder.Document.DocumentNumber;

                    // If the request has children, add them recursively
                    if (request.HasChildren())
                    {
                        foreach (RequestedProductionOrder child in request.Children)
                        {
                            AddRequestedProductionOrder(child, productionOrder);
                        }
                    }
                }
                else
                {
                    prodOrderId = 0;
                }
            }

            return pqUpdates;
        }

        /// <summary>
        /// Adds the standard production order.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <param name="parent">The parent.</param>
        private void AddStandardProductionOrder(RequestedStandardProductionOrder request, FortressProductionOrderAdapter parent = null)
        {
            // Attempt to create the standard production order
            using (var productionOrder = new FortressProductionOrderAdapter(this.Company))
            {
                productionOrder.Document.ItemNo = request.ItemCode;
                productionOrder.Document.DueDate = DateTime.Now;
                

                if (parent != null)
                {
                    productionOrder.Parent = parent.Document.AbsoluteEntry;
                    productionOrder.Revision = parent.Revision;
                    productionOrder.SalesOrderLineNumber = parent.SalesOrderLineNumber;
                    productionOrder.Document.ProductionOrderOrigin = BoProductionOrderOriginEnum.bopooManual;
                    productionOrder.Revision = request.DrawingVersion;
                }
                else
                {
                    productionOrder.Document.ProductionOrderOrigin = BoProductionOrderOriginEnum.bopooSalesOrder;
                    productionOrder.Document.ProductionOrderOriginEntry = this.Document.DocEntry;
                    productionOrder.SalesOrderLineNumber = request.SalesOrderLine;
                    productionOrder.Revision = request.DrawingVersion;
                }

                productionOrder.Add();

                if (productionOrder.Exists)
                {
                    // If the request has children, add them recursively
                    if (request.HasChildren())
                    {
                        foreach (RequestedProductionOrder child in request.Children)
                        {
                            AddRequestedProductionOrder(child, productionOrder);
                        }
                    }
                }
            }            
        }

        private void AddAddBomProductionOrder(RequestedAddBomProductionOrder request)
        {
            // Attempt to create the standard production order
            using (var productionOrder = new FortressProductionOrderAdapter(this.Company))
            {
                productionOrder.Document.ProductionOrderType = BoProductionOrderTypeEnum.bopotSpecial;
                productionOrder.Document.ItemNo = request.ItemCode;
                productionOrder.Document.DueDate = DateTime.Now;
                productionOrder.Document.PlannedQuantity = request.Quantity;

                productionOrder.Document.ProductionOrderOrigin = BoProductionOrderOriginEnum.bopooSalesOrder;
                productionOrder.Document.ProductionOrderOriginEntry = this.Document.DocEntry;
                productionOrder.SalesOrderLineNumber = request.SalesOrderLine;
                productionOrder.Revision = request.DrawingVersion;

                ProductionOrders_Lines lines = productionOrder.Document.Lines;

                lines.SetCurrentLine(lines.Count - 1);

                lines.ItemNo = "AddBOMToSAP";
                lines.BaseQuantity = request.Quantity;
                lines.PlannedQuantity = 1;
                lines.Add();

                productionOrder.Add();

                //if (productionOrder.Exists)
                //{
                //    // If the request has children, add them recursively
                //    if (request.HasChildren())
                //    {
                //        foreach (RequestedProductionOrder child in request.Children)
                //        {
                //            AddRequestedProductionOrder(child, productionOrder);
                //        }
                //    }
                //}
            }            
        }
        

        /// <summary>
        /// Creates the production order for the given line.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">The item code</param>
        /// <returns>A RequestedProductionOrder containing information about the Production order that needs 
        /// to be Added</returns>
        private RequestedProductionOrder PrepareRequest(int lineNum, string itemCode, double quantity)
        {
            int drawingVersion = this.GetCurrentLineRevision(lineNum);

            if (this.IsStandardBOM(itemCode))
            {
                return new RequestedStandardProductionOrder(this.Document.DocEntry, lineNum, itemCode, drawingVersion);
            }
            else
            {
                return new RequestedAddBomProductionOrder(this.Document.DocEntry, lineNum, itemCode, drawingVersion, quantity);
            }
        }

        /// <summary>
        /// Creates the production orders for the given line.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">The item code</param>
        /// <returns>A RequestedProductionOrder containing information about the Production order that needs 
        /// to be Added</returns>
        //[Obsolete("DoPrepareProductionOrders is deprecated, please use DoPrepareProductionOrder instead.")]
        //private RequestedProductionOrder DoPrepareProductionOrders(int lineNum, string itemCode)
        //{
        //    RequestedProductionOrder rpo;

        //    bool xmlFound = false;
        //    int drawingVersion = this.GetCurrentLineRevision(lineNum);

        //    // 1. Check if XML file exists in SalesOrder/LineNum folder
        //    // Note: XML Files can only exist when QTY = 1 (Configurations can only exist in this case)
        //    if ((int)this.GetLineQuantity(lineNum) == 1 && this.CheckIfXmlExists(lineNum, itemCode, drawingVersion))
        //    {
        //        // XML found
        //        xmlFound = true;
        //    }
        //    else
        //    {
        //        // Check for Standard BOM
        //        if (!this.IsStandardBOM(itemCode))
        //        {
        //            // Can't create Production Order!
        //            throw new FortressException(string.Format("Cannot create Production Order for line {0}, item {1}. No XML found, and line item is not a standard BOM", lineNum, itemCode));
        //        }
        //    }

        //    // 2. Look for Sub Assemblies
        //    if (xmlFound)
        //    {
        //        rpo = this.DoPrepareProductionOrderFromFile(lineNum, itemCode, drawingVersion);
        //    }
        //    else
        //    {
        //        rpo = this.DoPrepareProductionOrderFromBom(lineNum, itemCode, drawingVersion);
        //    }

        //    return rpo;
        //}



        /// <summary>
        /// Creates a production order from a Standard BOM.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">The item code.</param>
        /// <returns>A RequestedProductionOrder containing information about the Production order that needs 
        /// to be Added</returns>
        //private RequestedStandardProductionOrder DoPrepareProductionOrderFromBom(int lineNum, string itemCode, int drawingVersion)
        //{
        //    var rspo = new RequestedStandardProductionOrder(this.Document.DocEntry, lineNum, itemCode, drawingVersion);

        //    // Check if there are any sub assemblies
        //    IEnumerable<string> subAssemblies = this.GetSubAssemblies(lineNum, itemCode);

        //    // We need to create a production order for each one of these
        //    foreach (string subAssy in subAssemblies)
        //    {
        //        // Recursively call the DoCreateProductionOrders on the sub-assembly
        //        RequestedProductionOrder child = this.DoPrepareProductionOrders(lineNum, subAssy);
        //        rspo.AddChild(child);
        //    }

        //    return rspo;
        //}

        /// <summary>
        /// Creates a production order from a file.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">The item code.</param>
        /// <returns></returns>
        //private RequestedSpecialProductionOrder DoPrepareProductionOrderFromFile(int lineNum, string itemCode, int drawingVersion)
        //{
        //    // Get the full path to the Sales Order
        //    var slLineDir = this.GetSalesOrderLineDirectory(lineNum, itemCode);
        //    if (slLineDir == null) return null;
        //    string fileName = slLineDir + "/" + this.GetSalesOrderLineXmlFileName(lineNum, itemCode);

        //    var rspo = new RequestedSpecialProductionOrder(this.Document.DocEntry, lineNum, fileName, drawingVersion);

        //    ProductionOrders po;
        //    string errorMsg = string.Empty;

        //    try
        //    {
        //        XmlDocument xDoc = new XmlDocument();
        //        xDoc.Load(fileName);

        //        XmlNodeList nl = xDoc.SelectNodes("//ItemNo");

        //        foreach (XmlNode node in nl)
        //        {
        //            if (node.InnerText.Trim().Length > 20)
        //            {
        //                errorMsg = "[Item " + node.InnerText + "]";
        //            }
        //        }

        //        // Load the Production Order from an XML (BOM) file
        //        po = (ProductionOrders)this.Company.GetBusinessObjectFromXML(fileName, 0);
        //    }
        //    catch (Exception ex)
        //    {
        //        // Re-throw with added info 
        //        throw new FortressException(string.Format("Unable to load XML. " + errorMsg + "{0}", ex.Message), ex);
        //    }

        //    using (var poAdapter = new FortressProductionOrderAdapter(this.Company, po))
        //    {
        //        // Check every line for a sub-assembly
        //        foreach (ProductionOrders_Lines line in poAdapter.Lines)
        //        {
        //            using (var itemAdapter = new ItemAdapter(this.Company, line.ItemNo))
        //            {
        //                if (string.IsNullOrEmpty(line.ItemNo))
        //                {
        //                    throw new FortressException(string.Format("Line {0} does not have an Item Number on Production Order XML in file {1}", lineNum, fileName));
        //                }

        //                // Check if this is a sub-assembly (XML Special PO, or Standard BOM)
        //                if (this.CheckIfXmlExists(lineNum, itemAdapter.ItemCode, this.GetCurrentLineRevision(lineNum))) // < 110429 - Dom : Don't create PrO for standard BOMs />   || itemAdapter.IsBOM())
        //                {
        //                    // Recursively call DoCreateProductionOrders on the sub assembly
        //                    RequestedProductionOrder child = this.DoPrepareProductionOrders(lineNum, line.ItemNo);
        //                    rspo.AddChild(child);
        //                }
        //            }
        //        }
        //    }

        //    return rspo;
        //}

        /// <summary>
        /// Gets the kitted items for the given item code.
        /// </summary>
        /// <remarks>
        /// This does a deep search for BOMs.
        /// </remarks>
        /// <param name="lineNum">the line number</param>
        /// <param name="itemCode">The item code.</param>
        /// <returns>A list of item codes which are sub kits</returns>
        private IEnumerable<string> GetSubAssemblies(int lineNum, string itemCode)
        {
            IList<string> kits = new List<string>();

            using (var itemAdapter = new ItemAdapter(this.Company, itemCode))
            {
                if (itemAdapter.IsBOM())
                {
                    // Check if any of the components are assemblies (Special or Standard)
                    foreach (string component in itemAdapter.GetKittedItems())
                    {
                        using (var componentAdapter = new ItemAdapter(this.Company, component))
                        {
                            // Check if the subcomponent is a sub assembly BOM
                            if (this.CheckIfXmlExists(lineNum, componentAdapter.ItemCode, this.GetCurrentLineRevision(lineNum)) || componentAdapter.IsBOM())
                            {
                                kits.Add(component);
                            }

                            // Recursively get any children sub assemblies, and include them in the list
                            IEnumerable<string> subKits = GetSubAssemblies(lineNum, component);
                            foreach (string subKit in subKits)
                            {
                                kits.Add(subKit);
                            }
                        }
                    }
                }
            }

            return kits;
        }

        /// <summary>
        /// Sets the production order line status.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="pos">The production order status to set</param>
        private void SetProductionOrderLineStatus(int lineNum, ProductionOrderStatus pos)
        {
            this.SetLine(lineNum);
            this.SetLineUserDefinedField("U_XX_ProdStatus", ProductionOrderStatusHelper.GetStringValue(pos));
        }

        /// <summary>
        /// Gets the line status.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns></returns>
        private ApprovalStatus GetLineApprovalStatus(int lineNum)
        {
            try
            {
                this.SetLine(lineNum);
                return ApprovalStatusHelper.GetEnumFromString(this.GetLineUserDefinedField("U_XX_ApprovalReq").ToString());
            }
            catch
            {
                return ApprovalStatus.None;
            }
        }

        /// <summary>
        /// Gets the line approval statuses.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns></returns>
        private IList<string> GetLineApprovalStatuses(int lineNum)
        {
            IList<string> statuses = new List<string>();

            string query = string.Format(
                @"select * from [@XX_APPROV] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}", this.Document.DocEntry,
                lineNum);

            using (var recordset = new RecordsetAdapter(this.Company, query))
            {
                while (!recordset.EoF)
                {
                    statuses.Add(recordset.FieldValue("U_XX_Status").ToString());
                    recordset.MoveNext();
                }
            }

            return statuses;
        }

        /// <summary>
        /// Determines whether the configuration (if any) for the given line has passed.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns>
        /// 	<c>true</c> if the configuration for the given line has passed; otherwise, <c>false</c>.
        /// </returns>
        private bool HasConfigurationPassed(int lineNum)
        {
            bool hasPassed = false;

            int revision = this.GetCurrentLineRevision(lineNum);

            string query = string.Format(@"
                    select ISNULL(U_XX_Status, '') as LineStatus FROM
                    (
	                    select ISNULL(U_XX_Status, '') as U_XX_Status FROM [@XX_CONVBLT] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1} AND U_XX_Revis = '{2}' AND U_XX_Status = 'Passed'
	                    union all
	                    select ISNULL(U_XX_Status, '') as U_XX_Status FROM [@XX_METDET] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1} AND U_XX_Revis = '{2}' AND U_XX_Status = 'Passed'
                    ) as revisions", this.DocEntry, lineNum, revision);

            using (var recordset = new RecordsetAdapter(this.Company, query))
            {
                if (!recordset.EoF)
                {
                    hasPassed = true;
                }
            }

            return hasPassed;
        }

        #region Private Helper Method(s)

        /// <summary>
        /// Applies the pattern.
        /// </summary>
        /// <param name="pattern">The pattern.</param>
        /// <returns></returns>
        public string ApplyPattern(string pattern)
        {
            string applied = pattern;

            applied = ReplacePattern(applied, "DocEntry", this.DocEntry);
            applied = ReplacePattern(applied, "DocNumber", this.DocNumber);
            applied = ReplacePattern(applied, "DocDate", this.Document.DocDate);
            applied = ReplacePattern(applied, "DocDueDate", this.Document.DocDueDate);

            return applied;
        }        

        /// <summary>
        /// Applies the pattern.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="itemCode">the item code</param>
        /// <param name="pattern">The pattern.</param>
        /// <returns>the string value with pattern applied</returns>
        public string ApplyPattern(int lineNum, string itemCode, string pattern)
        {
            string applied = pattern;

            applied = ReplacePattern(applied, "LineNum", lineNum);
            applied = ReplacePattern(applied, "ItemCode", itemCode);
            applied = ReplacePattern(applied, "Revision", this.GetCurrentLineRevision(lineNum));

            return applied;
        }

        /// <summary>
        /// Applies the pattern.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <param name="revision">the revision</param>
        /// <param name="pattern">The pattern.</param>
        /// <param name="itemCode">the item code</param>
        /// <returns>the string value with pattern applied</returns>
        public string ApplyPattern(int lineNum, string itemCode, int revision, string pattern)
        {
            string applied = pattern;

            applied = ReplacePattern(applied, "LineNum", lineNum);
            applied = ReplacePattern(applied, "ItemCode", itemCode);
            applied = ReplacePattern(applied, "Revision", revision);

            return applied;
        }

        #endregion Private Helper Method(s)

        #endregion Private Method(s)

        #region Static Method(s)

        public static DirectoryInfo GetOrCreateSalesOrderDirectory(Company company, int docEntry)
        {
            DirectoryInfo dirInfo = FortressConfigurationHelper.GetDrawingDirectory(company);

            int startingRange = FortressConfigurationHelper.GetSalesOrderDirectoryInitial(company);
            int factor = FortressConfigurationHelper.GetSalesOrderDirectoryFactor(company);
            string format = FortressConfigurationHelper.GetSalesOrderDirectoryNumericFormat(company);

            using (var soAdapter = new FortressSalesOrderAdapter(company, docEntry))
            {
                DirectoryInfo soDir = soAdapter.CreateDirectory(soAdapter.DocNumber, startingRange, factor, dirInfo, format, company);
                return soDir;
            }
        }

        private DirectoryInfo CreateDirectory(int actualNumber, int range, int factor, DirectoryInfo directory, string format, Company company)
        {
            int start = (actualNumber / range) * range;
            int end = start + range - 1;

            string dirName = start.ToString(format) + "-" + end.ToString(format);

            DirectoryInfo foundDir = null;
            DirectoryInfo[] dirs = directory.GetDirectories(dirName);
            if (dirs.Length == 0)
            {
                foundDir = directory.CreateSubdirectory(dirName);
            }
            else
            {
                foundDir = dirs[0];
            }

            DirectoryInfo dInfo = new DirectoryInfo(foundDir.FullName);
            if (range > factor)
            {
                return CreateDirectory(actualNumber, range / factor, factor, dInfo, format, company);
            }

            
            // Check if the actual SO directory exists
            string pattern = FortressConfigurationHelper.GetSalesOrderDirectoryPattern(company);
            
            string soDirName = this.ApplyPattern(pattern);

            DirectoryInfo[] soDirs = dInfo.GetDirectories(soDirName);
            if (soDirs.Length > 0)
            {
                dInfo = soDirs[0];
            }
            else
            {
                dInfo = dInfo.CreateSubdirectory(soDirName);
            }
            dInfo = new DirectoryInfo(dInfo.FullName);
            return dInfo;
        }

        /// <summary>
        /// Replaces the pattern.
        /// </summary>
        /// <param name="toReplace">To replace.</param>
        /// <param name="patternName">Name of the pattern.</param>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        private static string ReplacePattern(string toReplace, string patternName, string value)
        {
            string toReturn = toReplace;
            string param = string.Empty;

            if (toReturn.Contains("{" + patternName))
            {
                int startIndex = toReturn.IndexOf("{" + patternName);

                param = toReturn.Substring(startIndex);
                param = param.Substring(0, param.IndexOf('}') + 1);
            }

            if (string.IsNullOrEmpty(param))
                return toReturn;

            return toReturn.Replace(param, value);
        }

        /// <summary>
        /// Replaces the pattern.
        /// </summary>
        /// <param name="toReplace">To replace.</param>
        /// <param name="patternName">Name of the pattern.</param>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        private static string ReplacePattern(string toReplace, string patternName, int value)
        {
            string toReturn = toReplace;
            string param = string.Empty;

            string replacement = value.ToString();

            if (toReturn.Contains("{" + patternName))
            {
                int startIndex = toReturn.IndexOf("{" + patternName);

                param = toReturn.Substring(startIndex);
                param = param.Substring(0, param.IndexOf('}') + 1);

                if (param.Contains(':'))
                {
                    string format = param.Substring(param.IndexOf(':') + 1, (param.Length - param.IndexOf(':')) - 2);
                    replacement = value.ToString(format);
                }
            }

            if (string.IsNullOrEmpty(param))
                return toReturn;

            return toReturn.Replace(param, replacement);
        }

        /// <summary>
        /// Replaces the pattern.
        /// </summary>
        /// <param name="toReplace">To replace.</param>
        /// <param name="patternName">Name of the pattern.</param>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        private static string ReplacePattern(string toReplace, string patternName, DateTime value)
        {
            string toReturn = toReplace;
            string param = string.Empty;

            string replacement = value.ToString();

            if (toReturn.Contains("{" + patternName))
            {
                int startIndex = toReturn.IndexOf("{" + patternName);

                param = toReturn.Substring(startIndex);
                param = param.Substring(0, param.IndexOf('}') + 1);

                if (param.Contains(':'))
                {
                    string format = param.Substring(param.IndexOf(':') + 1, (param.Length - param.IndexOf(':')) - 2);
                    replacement = value.ToString(format);
                }
            }

            if (string.IsNullOrEmpty(param))
                return toReturn;

            return toReturn.Replace(param, replacement);
        }

        #endregion Static Method(s)

        #endregion Method(s)

    }
}
