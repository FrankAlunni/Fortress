//-----------------------------------------------------------------------
// <copyright file="Commission.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
// <email>frank.alunni@b1computing.com</email>
//-----------------------------------------------------------------------

namespace B1C.SAP.Addons.Configurator.Forms
{
    #region Usings

    using System;
    using B1C.SAP.UI.Model.EventArguments;
    using B1C.SAP.DI.Helpers;
    using Fortress.DI.Helpers;
    using B1C.SAP.DI.BusinessAdapters;
    using B1C.SAP.DI.Business;
    using Fortress.DI.Model.Exceptions;
    using SAPbobsCOM;
    using System.Collections.Generic;
    using B1C.SAP.DI.BusinessAdapters.Administration;
    using System.Web;

    #endregion Usings

    public class Commission : Window.FortressForm
    {
        #region Constuctors

        /// <summary>
        /// Initializes a new instance of the <see cref="Commission"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        public Commission(UI.AddOn addOn, string xmlPath, string formType)
            : base(addOn, xmlPath, formType)
        {
            this.ModifiableLines = new Dictionary<string, bool>();
        }

        #endregion Constuctors

        enum CommissionType
        {
            NotFound,
            ConveyorBelt,
            MetalDetector,
            Pipeline,
            TableSystem,
            GravityRejectSystem
        }

        private bool _Validating = false;
        private double _LineTotal = 0;


        private double CommissionAvailable
        {
            get
            {
                double total;
                double.TryParse(this.Field("amDS").Value, out total);
                return total;
            }

            set
            {
                if (this.Field("amDS").Value != value.ToString())
                {
                    this.Field("amDS").Value = value.ToString();
                }
            }
        }

        private double TotalCommission
        {
            get
            {
                double total;
                double.TryParse(this.Field("totDS").Value, out total);
                return total;
            }

            set
            {
                if (this.Field("totDS").Value != value.ToString())
                {
                    this.Field("totDS").Value = value.ToString();
                }
            }
        }

        private double CommissionAvailablePct
        {
            get
            {
                double total;
                double.TryParse(this.Field("cmDS").Value, out total);
                return total;
            }

            set
            {
                
                if (this.Field("cmDS").Value != value.ToString())
                {
                    this.Field("cmDS").Value = value.ToString();
                }
            }
        }

        #region Methods

        private bool IsServiceRow(int row)
        {
            return this.ControlManager.Matrix("mtx_0")[row, "col_2"] == "Service";
        }

        private double RowRatePct(int row)
        { 
            double rate = 0;
            double.TryParse(this.ControlManager.Matrix("mtx_0")[row, "col_7"], out rate);

            return rate;
        }

        private double RowAmount(int row)
        {
            double amount = 0;
            double.TryParse(this.ControlManager.Matrix("mtx_0")[row, "col_19"], out amount);

            return amount;
        }


        /// <summary>
        /// Initializes this instance.
        /// </summary>
        public override void Initialize()
        {
            this.ButtonClick += this.OnButtonClick;
            this.AfterArgumentsSet += this.OnAfterArgumentsSet;
            this.FormResize += this.OnFormResize;
            this.AfterValidate += this.OnAfterValidate;
            this.AfterComboSelect += this.OnAfterValidate;
            this.BeforeComboSelect += this.OnBeforeComboSelect;
            this.ChooseFromList += this.OnChooseFromList;
            this.AfterUpdate += this.OnAfterUpdate;
            base.Initialize();
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public override void Dispose()
        {
            this.ButtonClick -= this.OnButtonClick;
            this.AfterArgumentsSet -= this.OnAfterArgumentsSet;
            this.FormResize -= this.OnFormResize;
            this.AfterValidate -= this.OnAfterValidate;
            this.AfterComboSelect -= this.OnAfterValidate;
            this.BeforeComboSelect -= this.OnBeforeComboSelect;
            this.ChooseFromList -= this.OnChooseFromList;
            this.AfterUpdate -= this.OnAfterUpdate;
            base.Dispose();
        }

        private IDictionary<string, bool> ModifiableLines { get; set; }

        private void RowUpdated(bool updated, int row)
        {
            if (updated)
            {
                this.ControlManager.Matrix("mtx_0")[row, "col_5"] = "Y";
            }
            else
            {
                this.ControlManager.Matrix("mtx_0")[row, "col_5"] = "N";
            }
        }

        public void OnAfterUpdate(object sender, FormEventArgs e)
        {
            this.UpdateModifiableLines();
        }

        private void OnChooseFromList(object sender, ChooseFromListEventArgs e)
        {
            var dt = e.SelectedObjects;
            var matrix = this.ControlManager.Matrix("mtx_0");

            int sr = matrix.SelectedRow();
            int row = e.SelectedRow;
            SAPbouiCOM.EditText edit = (SAPbouiCOM.EditText)matrix.Columns.Item("col_6").Cells.Item(row).Specific;

            try
            {
                this.SapForm.Select();
                this.SapForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                edit.Value = dt.GetValue(1, 0).ToString();
            }
            catch { }

            // Mark the row as updated
            RowUpdated(true, row);
        }

        /// <summary>
        /// Called when [before key down].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.ItemEventArgs"/> instance containing the event data.</param>
        private void OnBeforeComboSelect(object sender, ItemEventArgs e)
        {
            if (!this.IsAuthorized("Edit"))
            {
                var matrix = this.ControlManager.Matrix("mtx_0");

                // If is new record can be edited
                if (matrix[e.Row, "col_1"] != "<New>" || (matrix[e.Row, "col_16"] == "Approved" && !this.IsAuthorized("EditApr")))
                {
                    e.CancelEvent = true;
                    e.CancelMessage = "User not autorized to perform this operation on an existing record.";
                }
            }
        }


        private void CalculateRowAmount(int row, string colId)
        {
            RowUpdated(true, row);

            double total = 0;
            if (IsServiceRow(row))
            {
                total = _LineTotal;
            }
            else
            {
                total = CommissionAvailable;
            }

            if (colId == "col_7")
            {
                // Update Amount
                double rate = RowRatePct(row);
                double amount = (rate == 0) ? 0 : ((total * rate) / 100);

                this.ControlManager.Matrix("mtx_0")[row, "col_19"] = amount.ToString();
            }
            else if (colId == "col_19")
            {
                // Update Rate
                double amount = RowAmount(row);
                double rate = (total == 0) ? 0 : ((amount / total) * 100);

                this.ControlManager.Matrix("mtx_0")[row, "col_7"] = rate.ToString();
            }

        }

        private void OnAfterValidate(object sender, ItemEventArgs e)
        {
            if (e.ItemId == "mtx_0")
            {
                try
                {
                    if (_Validating) return;

                    _Validating = true;
                    CalculateRowAmount(e.Row, e.ColId);
                }
                finally
                {
                    _Validating = false;
                }
            }
        }

        /// <summary>
        /// Called when [form resize].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        private void OnFormResize(object sender, FormEventArgs e)
        {
            var matrix = this.ControlManager.Matrix("mtx_0");

            matrix.Columns.Item(0).Width = 20;
            matrix.Columns.Item(1).Width = 0;
            matrix.Columns.Item(2).Width = 70;
            matrix.Columns.Item(4).Width = 50;
            matrix.Columns.Item(5).Width = 60;
            matrix.Columns.Item(6).Width = 60;
            matrix.Columns.Item(7).Width = 40;
            matrix.Columns.Item(8).Width = 0;
            matrix.Columns.Item(9).Width = 70;
            matrix.Columns.Item(10).Width = 60;
            matrix.Columns.Item(11).Width = 40;
            matrix.FillWidthOnColumn(3);
        }

        private void UpdateModifiableLines()
        {
            this.ModifiableLines = new Dictionary<string, bool>();
            using (var rec = new RecordsetAdapter(this.AddOn.Company, string.Format("SELECT Code, U_XX_Approved FROM [@XX_COMMISS] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}", this.OrderId, this.OrderLine)))
            {
                while (!rec.EoF)
                {
                    string code = rec.FieldValue("Code").ToString();
                    string app = rec.FieldValue("U_XX_Approved").ToString();
                    this.ModifiableLines[code] = app != "Approved";
                    rec.MoveNext();
                }
            }
        }

        private void LoadMatrix()
        {

            var statusTable = this.DataSource.UserTable("CommissDT");
            statusTable.Rows.Clear();

            string query = string.Format(@"SELECT DocEntry, Code, U_XX_Type as ComType, Name, U_XX_Rate as Rate,  ISNULL(U_XX_Amount, 0) as Amount,  CONVERT(varchar(10), CreateDate, 121) as CreateDate, LEFT(CreateTime, len(CreateTime)-2) + ':' + RIGHT(CreateTime, 2) as CreateTime, CONVERT(varchar(10), UpdateDate, 121) as UpdateDate, LEFT(UpdateTime, len(UpdateTime)-2) + ':' + RIGHT(UpdateTime, 2) as UpdateTime, U_XX_Approved  as Approved FROM [@XX_COMMISS] WITH (NOLOCK) WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}", this.OrderId, this.OrderLine);

            statusTable.ExecuteQuery(query);

            this.ControlManager.Matrix("mtx_0").BaseObject.LoadFromDataSource();
        }

        private double GetDefaultCommission(string itemCode, out CommissionType currentCommission)
        {
            string prefixConvBelt;
            ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "ConvPref", out prefixConvBelt);
            string prefixMetalDet;
            ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "DetPref", out prefixMetalDet);
            string prefixPipeline;
            ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "PpPref", out prefixPipeline);
            string prefixTablet;
            ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "TbtPref", out prefixTablet);
            string prefixGravity;
            ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "GrvPref", out prefixGravity);

            currentCommission = CommissionType.NotFound;
            if (itemCode.StartsWith(prefixConvBelt))
            {
                currentCommission = CommissionType.ConveyorBelt;
            }
            else if (itemCode.StartsWith(prefixMetalDet))
            {
                currentCommission = CommissionType.MetalDetector;
            }
            else if (itemCode.StartsWith(prefixPipeline))
            {
                currentCommission = CommissionType.Pipeline;
            }
            else if (itemCode.StartsWith(prefixTablet))
            {
                currentCommission = CommissionType.TableSystem;
            }
            else if (itemCode.StartsWith(prefixGravity))
            {
                currentCommission = CommissionType.GravityRejectSystem;
            }

            switch (currentCommission)
            {
                case CommissionType.ConveyorBelt:
                    return FortressConfigurationHelper.GetConveyorCommisionRate(this.AddOn.Company);
                case CommissionType.GravityRejectSystem:
                    return FortressConfigurationHelper.GetGravityRejectSystemCommisionRate(this.AddOn.Company);
                case CommissionType.MetalDetector:
                    return FortressConfigurationHelper.GetMetalDetectorCommisionRate(this.AddOn.Company);
                case CommissionType.Pipeline:
                    return FortressConfigurationHelper.GetPipelineCommisionRate(this.AddOn.Company);
                case CommissionType.TableSystem:
                    return FortressConfigurationHelper.GetTabletSystemCommisionRate(this.AddOn.Company);
                default:
                    return 0;
            }
        }

        /// <summary>
        /// Called when [after arguments set].
        /// </summary>
        /// <param name="sender">The sender.</param>
        private void OnAfterArgumentsSet(object sender)
        {
            this.UpdateModifiableLines();

            var currentUser = UserAdapter.Current(this.AddOn.Company);
            string itemCode = string.Empty;

            if (this.IsNewRecord)
            {
                string itemCodeQuery = string.Format("SELECT ItemCode, CASE WHEN TotalFrgn > 0 Then TotalFrgn ELSE LineTotal END as LineTotal FROM RDR1 WHERE DocEntry = {0} AND LineNum = {1}", this.OrderId, this.OrderLine);
                using (var rec = new RecordsetAdapter(this.AddOn.Company, itemCodeQuery))
                {
                    if (!rec.EoF)
                    {
                        _LineTotal = (double)rec.FieldValue("LineTotal");
                        itemCode = rec.FieldValue("ItemCode").ToString();
                    }
                }

                double rate = 100;
                double total = CommissionAvailable;
                //double.TryParse(this.Field("amDS").Value, out total);

                double amount = (rate == 0) ? 0 : ((total * rate) / 100);


                // Get the default commision pct.
                CommissionType currentCommission = CommissionType.NotFound;
                double dfltCommis = this.GetDefaultCommission(itemCode, out currentCommission);

                string dftlSP = string.Empty;
                using (var rc = new RecordsetAdapter(this.AddOn.Company, "select OSLP.SlpName from ORDR inner join OCRD  on ORDR.CardCode = OCRD.CardCode inner join OSLP on OCRD.SlpCode = OSLP.SlpCode  WHERE ORDR.DocEntry = " + this.OrderId))
                {
                    if (!rc.EoF)
                    {
                        dftlSP = rc.FieldValue("SlpName").ToString();
                    }
                }

                CommissionAvailablePct = dfltCommis;
                CommissionAvailable = (_LineTotal * (dfltCommis / 100));
                TotalCommission = CommissionAvailable;


                var loadData = new System.Text.StringBuilder();
                loadData.AppendLine("<DataLoader UDOName='Commission' UDOType='MasterData'>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", new TableFieldAdapter(this.AddOn.Company, "@XX_COMMISS", "Code").NextCode("CMM", string.Empty, 0));
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(dftlSP));
                loadData.AppendFormat("<Field Name='U_XX_Rate' Type='System.Double' Value='{0}' />", rate);
                loadData.AppendFormat("<Field Name='U_XX_Amount' Type='System.Double' Value='{0}' />", amount);
                loadData.AppendFormat("<Field Name='U_XX_Type' Type='System.String' Value='{0}' />", "Full");
                loadData.AppendFormat("<Field Name='U_XX_OrderNo' Type='System.Int32' Value='{0}' />", this.OrderId);
                loadData.AppendFormat("<Field Name='U_XX_OrdrLnNo' Type='System.Int32' Value='{0}' />", this.OrderLine);
                loadData.AppendFormat("<Field Name='U_XX_CommisAv' Type='System.Double' Value='{0}' />", dfltCommis);
                loadData.AppendLine("</Record>");
                loadData.AppendLine("</DataLoader>");

                DataLoader loader = DataLoader.InstanceFromXMLString(loadData.ToString());
                if (loader != null)
                {
                    loader.Process(this.AddOn.Company);

                    if (!string.IsNullOrEmpty(loader.ErrorMessage))
                    {
                        throw new Exception("Data loader exception: " + loader.ErrorMessage);
                    }
                }

                //this.ControlManager.EditText("soCommis").Value = dfltCommis.ToString();
                //this.ControlManager.EditText("soComAmt").Value = (lineTotal * dfltCommis).ToString();
            }
            else
            {
                UpdateCommissionAmount();

                //double commisAv = 0;
                //string q = string.Format("SELECT TOP 1 U_XX_CommisAv FROM [@XX_COMMISS] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}", this.OrderId, this.OrderLine);
                //using (var rec = new RecordsetAdapter(this.AddOn.Company, q))
                //{
                //    if (!rec.EoF)
                //    {
                //        commisAv = (double)rec.FieldValue("U_XX_CommisAv");
                //    }
                //}

                //CommissionAvailablePct = commisAv;
                ////this.ControlManager.EditText("soCommis").Value = commisAv.ToString();

                //double totalRate = 0;
                //string amntQuery = string.Format("SELECT Sum(U_XX_Rate) as totalRate FROM [@XX_COMMISS] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1} AND U_XX_Type <> 'Service'", this.OrderId, this.OrderLine);
                //using (var rec = new RecordsetAdapter(this.AddOn.Company, amntQuery))
                //{
                //    if (!rec.EoF)
                //    {
                //        totalRate = (double)rec.FieldValue("totalRate");
                //    }
                //}

                //double serviceRate = 0;
                //string svcAmntQuery = string.Format("SELECT Sum(U_XX_Rate) as totalRate FROM [@XX_COMMISS] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1} AND U_XX_Type = 'Service'", this.OrderId, this.OrderLine);
                //using (var rec = new RecordsetAdapter(this.AddOn.Company, svcAmntQuery))
                //{
                //    if (!rec.EoF)
                //    {
                //        serviceRate += (double)rec.FieldValue("totalRate");
                //    }
                //}


                //CommissionAvailable = ((lineTotal * (commisAv / 100) * (totalRate / 100)) + (lineTotal * (serviceRate / 100)));
                ////this.ControlManager.EditText("soComAmt").Value = (lineTotal * commisAv * totalRate).ToString();
            }

            this.ControlManager.EditText("soNumber").Value = this.OrderId;
            this.Field("soDocNoDS").Value = this.DocumentNumber;
            this.ControlManager.EditText("soLineNum").Value = this.OrderLine;
            this.ControlManager.EditText("soRowNum").Value = this.RowNumber;

            this.SapForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

            AddChooseFromList();
            var matrix = this.ControlManager.Matrix("mtx_0");
            var cols = matrix.Columns;
            var colmn = cols.Item("col_6");
            colmn.ChooseFromListUID = "UDCFL1";
            colmn.ChooseFromListAlias = "SlpName";

            // Load Datatable and Matrix
            this.LoadMatrix();
            if (!this.IsAuthorized("Add"))
            {
                this.ControlManager.Button("3").Visible = false;
            }

            if (!this.IsAuthorized("Delete"))
            {
                this.ControlManager.Button("delRow").Visible = false;
            }

            CalculateRowAmount(1, "col_7");

        }


        private void AddChooseFromList()
        {
            // Choose From List Collection object to store the collection
            SAPbouiCOM.ChooseFromListCollection oCFLs;

            // Getting the Form Choose From Lists 
            oCFLs = this.SapForm.ChooseFromLists;

            // Choose From List object to store the ChooseFromList
            SAPbouiCOM.ChooseFromList oCFL;

            // ChooseFromListCreationParams to create the parameters for the CFL
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;
            oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)this.AddOn.Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "53";
            oCFLCreationParams.UniqueID = "UDCFL1";
            oCFL = oCFLs.Add(oCFLCreationParams);
        }

        /// <summary>
        /// Called when [button click].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.ItemEventArgs"/> instance containing the event data.</param>
        private void OnButtonClick(object sender, ItemEventArgs e)
        {
            switch (e.ItemId)
            {
                case "delRow":
                    this.DeleteRow();
                    break;
                case "3":
                    this.AddStatus();
                    break;
                case "1":
                    this.UpdateRecord();
                    break;
            }
        }

        private void DeleteRow()
        {

            var matrix = this.ControlManager.Matrix("mtx_0");

            if (matrix[matrix.SelectedRow(), "col_16"] == "Approved" && !IsAuthorized("DelApr"))
            {
                throw new FortressException("User is not authorized to delete approved row ");
            }


            if (matrix != null)
            {
                if (matrix.SelectedRow() != -1)
                {
                    string docEntry = matrix[matrix.SelectedRow(), "col_1"];

                    RecordsetAdapter.Execute(
                            this.AddOn.Company,
                            string.Format("DELETE FROM [@XX_COMMISS] WHERE Code = '{0}'", docEntry));
                    matrix.BaseObject.DeleteRow(matrix.SelectedRow());
                }
                else
                {
                    this.AddOn.StatusBar.ErrorMessage = "Select the row to be deleted.";
                }
            }
            else
            {
                this.AddOn.StatusBar.ErrorMessage = "Select the row to be deleted.";
            }
        }

        /// <summary>
        /// Adds the status.
        /// </summary>
        private void AddStatus()
        {
            try
            {
                double rate = 100 - this.TotalCommissionPercentage();

                rate = (rate < 0) ? 0 : rate;

                double total = CommissionAvailable;
                //double.TryParse(this.Field("amDS").Value, out total);

                double amount = (rate == 0) ? 0 : ((total * rate) / 100);

                var selectedSerial = this.DataSource.UserTable("CommissDT");

                if (selectedSerial != null)
                {
                    // Check the value of the last row
                    var lastCode = selectedSerial.GetValue("Code", selectedSerial.Rows.Count - 1);
                    if (!string.IsNullOrEmpty(lastCode.ToString()))
                    {
                        selectedSerial.Rows.Add(1);
                    }
                    selectedSerial.SetValue("DocEntry", selectedSerial.Rows.Count - 1, new TableFieldAdapter(this.AddOn.Company, "@XX_COMMISS", "DocEntry").NextCode(string.Empty, string.Empty));
                    selectedSerial.SetValue("Code", selectedSerial.Rows.Count - 1, "<New>");
                    selectedSerial.SetValue("ComType", selectedSerial.Rows.Count - 1, "Engineering");
                    selectedSerial.SetValue("Rate", selectedSerial.Rows.Count - 1, rate);
                    selectedSerial.SetValue("Amount", selectedSerial.Rows.Count - 1, amount);
                    selectedSerial.SetValue("Approved", selectedSerial.Rows.Count - 1, "None");
                    selectedSerial.SetValue("CreateDate", selectedSerial.Rows.Count - 1, DateTime.Now.ToString("yyyyMMdd"));
                    selectedSerial.SetValue("CreateTime", selectedSerial.Rows.Count - 1, DateTime.Now.ToString("hhmm"));
                }

                var matrix = this.ControlManager.Matrix("mtx_0");
                if (matrix != null)
                {
                    matrix.BaseObject.LoadFromDataSource();
                }

                this.EnabledButton("delRow", false);
                this.EnabledButton("3", false);

                this.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
                return;
            }
        }

        private void EnabledButton(string buttonId, bool enabled)
        {
            try
            {
                this.ControlManager.Button(buttonId).Enabled = enabled;
            }
            catch
            {
            }

        }

        private void UpdateCommissionAmount()
        {
            string itemCodeQuery = string.Format("SELECT ItemCode, CASE WHEN TotalFrgn > 0 Then TotalFrgn ELSE LineTotal END as LineTotal FROM RDR1 WHERE DocEntry = {0} AND LineNum = {1}", this.OrderId, this.OrderLine);
            using (var rec = new RecordsetAdapter(this.AddOn.Company, itemCodeQuery))
            {
                if (!rec.EoF)
                {
                    _LineTotal = (double)rec.FieldValue("LineTotal");
                }
            }

            double commisAv = 0;
            string q = string.Format("SELECT TOP 1 U_XX_CommisAv FROM [@XX_COMMISS] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}", this.OrderId, this.OrderLine);
            using (var rec = new RecordsetAdapter(this.AddOn.Company, q))
            {
                if (!rec.EoF)
                {
                    commisAv = (double)rec.FieldValue("U_XX_CommisAv");
                }
            }

            CommissionAvailablePct = commisAv;
            CommissionAvailable = (_LineTotal * (CommissionAvailablePct / 100));

            //this.ControlManager.EditText("soCommis").Value = commisAv.ToString();

            double totalRate = 0;
            string amntQuery = string.Format("SELECT Sum(U_XX_Rate) as totalRate FROM [@XX_COMMISS] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1} AND U_XX_Type <> 'Service'", this.OrderId, this.OrderLine);
            using (var rec = new RecordsetAdapter(this.AddOn.Company, amntQuery))
            {
                if (!rec.EoF)
                {
                    totalRate = (double)rec.FieldValue("totalRate");
                }
            }

            double serviceRate = 0;
            string svcAmntQuery = string.Format("SELECT Sum(U_XX_Rate) as totalRate FROM [@XX_COMMISS] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1} AND U_XX_Type = 'Service'", this.OrderId, this.OrderLine);
            using (var rec = new RecordsetAdapter(this.AddOn.Company, svcAmntQuery))
            {
                if (!rec.EoF)
                {
                    serviceRate += (double)rec.FieldValue("totalRate");
                }
            }


            TotalCommission = ((_LineTotal * (commisAv / 100) * (totalRate / 100)) + (_LineTotal * (serviceRate / 100)));
        }

        private double TotalCommissionPercentage()
        {
            double totalCommissionPct = 0.0;
            var commissionMatrix = this.ControlManager.Matrix("mtx_0");

            for (int rowIndex = 1; rowIndex <= commissionMatrix.RowCount; rowIndex++)
            {
                if (!IsServiceRow(rowIndex))
                {
                    totalCommissionPct += RowRatePct(rowIndex);;
                }
            }

            return totalCommissionPct;
        }



        /// <summary>
        /// Updates the record.
        /// </summary>
        private void UpdateRecord()
        {
            if (this.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                return;
            }

            try
            {
                double commis;
                if (!double.TryParse(this.ControlManager.EditText("soCommis").Value, out commis))
                {
                    throw new FortressException("Total Commision must be a numeric value");
                }

                double totalCommissionPct = this.TotalCommissionPercentage();

                if (totalCommissionPct > 100)
                {
                    this.AddOn.MessageBox("Total commission allocated to each row adds up to more than 100%. Please modify your selelctions so that the total commission does not exceed 100%.");
                    return;
                }
                else if (totalCommissionPct < 100)
                {
                    this.AddOn.StatusBar.ErrorMessage = string.Format("{0}% of commission remaining to be paid out.", (100.0d - totalCommissionPct).ToString("#0.0000"));
                }

                var commissionMatrix = this.ControlManager.Matrix("mtx_0");
                if (commissionMatrix != null)
                {
                    // Check the value of the last row
                    for (int rowIndex = 1; rowIndex <= commissionMatrix.RowCount; rowIndex++)
                    {
                        var lastCode = commissionMatrix[rowIndex, "col_1"];
                        if (lastCode == "<New>")
                        {
                            // Append New records
                            var loadData = new System.Text.StringBuilder();
                            loadData.AppendLine("<DataLoader UDOName='Commission' UDOType='MasterData'>");
                            loadData.AppendLine("<Record>");
                            loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", new TableFieldAdapter(this.AddOn.Company, "@XX_COMMISS", "Code").NextCode("CMM", string.Empty));
                            loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(commissionMatrix[rowIndex, "col_6"]));
                            loadData.AppendFormat("<Field Name='U_XX_Rate' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(commissionMatrix[rowIndex, "col_7"]));
                            loadData.AppendFormat("<Field Name='U_XX_Amount' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(commissionMatrix[rowIndex, "col_19"]));
                            loadData.AppendFormat("<Field Name='U_XX_Type' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(commissionMatrix[rowIndex, "col_2"]));
                            loadData.AppendFormat("<Field Name='U_XX_OrderNo' Type='System.Int32' Value='{0}' />", this.OrderId);
                            loadData.AppendFormat("<Field Name='U_XX_OrdrLnNo' Type='System.Int32' Value='{0}' />", this.OrderLine);
                            loadData.AppendFormat("<Field Name='U_XX_CommisAv' Type='System.Double' Value='{0}' />", commis);
                            loadData.AppendFormat("<Field Name='U_XX_Approved' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(commissionMatrix[rowIndex, "col_16"]));
                            loadData.AppendLine("</Record>");
                            loadData.AppendLine("</DataLoader>");

                            DataLoader loader = DataLoader.InstanceFromXMLString(loadData.ToString());

                            if (loader != null)
                            {
                                loader.Process(this.AddOn.Company);

                                if (!string.IsNullOrEmpty(loader.ErrorMessage))
                                {
                                    throw new Exception("Data loader exception: " + loader.ErrorMessage);
                                }
                            }
                        }
                        else
                        {
                            // Check if record is updated
                            if (commissionMatrix[rowIndex, "col_5"] == "Y")
                            {
                                if (this.ModifiableLines.ContainsKey(lastCode) && !this.ModifiableLines[lastCode] && !this.IsAuthorized("Edit"))
                                {
                                    throw new FortressException("User is not authorized to update row " + rowIndex);
                                }
                                else if (this.ModifiableLines.ContainsKey(lastCode) && !this.ModifiableLines[lastCode] && !this.IsAuthorized("EditApr") && commissionMatrix[rowIndex, "col_16"] == "Approved")
                                {
                                    throw new FortressException("User is not authorized to update row " + rowIndex);
                                }

                                var statuses = this.AddOn.Company.UserTables.Item("XX_COMMISS");
                                statuses.GetByKey(lastCode);
                                statuses.UserFields.Fields.Item("U_XX_Type").Value = commissionMatrix[rowIndex, "col_2"];
                                statuses.Name = commissionMatrix[rowIndex, "col_6"];
                                statuses.UserFields.Fields.Item("U_XX_Rate").Value = commissionMatrix[rowIndex, "col_7"];
                                statuses.UserFields.Fields.Item("U_XX_Amount").Value = commissionMatrix[rowIndex, "col_19"];
                                statuses.UserFields.Fields.Item("U_XX_CommisAv").Value = commis;
                                statuses.UserFields.Fields.Item("U_XX_Approved").Value = commissionMatrix[rowIndex, "col_16"];
                                statuses.Update();
                            }
                            else
                            {
                                var statuses = this.AddOn.Company.UserTables.Item("XX_COMMISS");
                                statuses.GetByKey(lastCode);
                                double storedVal = (double)statuses.UserFields.Fields.Item("U_XX_CommisAv").Value;

                                if (storedVal != commis)
                                {
                                    statuses.UserFields.Fields.Item("U_XX_CommisAv").Value = commis;
                                    statuses.Update();
                                }
                            }
                        }
                    }
                }

                this.LoadMatrix();

                this.EnabledButton("delRow", true);
                this.EnabledButton("3", true);

                this.UpdateCommissionAmount();
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        #endregion Methods
    }
}
