//-----------------------------------------------------------------------
// <copyright file="SalesOrder.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
//-----------------------------------------------------------------------

namespace B1C.SAP.Addons.Configurator.Forms
{
    #region Using Directive(s)

    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Web;
    using B1C.SAP.Addons.Configurator.Window;
    using SAPbouiCOM;
    using B1C.SAP.UI.Model.EventArguments;
    using Fortress.DI.BusinessAdapters.Sales;
    using Fortress.DI.Model.Exceptions;
    using Fortress.DI.BusinessAdapters.Production;
    using Fortress.DI.Helpers;
    using B1C.SAP.DI.BusinessAdapters;
    using B1C.SAP.DI.Helpers;
    using B1C.SAP.DI.Business;

    #endregion Using Directive(s)

    public class SalesOrder : UI.Windows.Form
    {
        private bool _deleteInProgress;

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="UI.Windows.Form"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        public SalesOrder(UI.AddOn addOn)
            : base(addOn)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UI.Windows.Form"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="form">The SAP form.</param>
        public SalesOrder(UI.AddOn addOn, Form form)
            : base(addOn, form)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UI.Windows.Form"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        public SalesOrder(UI.AddOn addOn, string xmlPath, string formType)
            : base(addOn, xmlPath, formType)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UI.Windows.Form"/> class.
        /// </summary>
        /// <param name="addOn">The running add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        /// <param name="objectType">Type of the object.</param>
        public SalesOrder(UI.AddOn addOn, string xmlPath, string formType, string objectType)
            : base(addOn, xmlPath, formType, objectType)
        {
        }

        #endregion Constructor(s)

        #region Method(s)

        /// <summary>
        /// Initializes this instance.
        /// </summary>
        public override void Initialize()
        {
            base.Initialize();
            FormResize += this.OnFormResize;
            ButtonClick += this.OnButtonClick;
            AfterFormLoad += this.OnAfterFormLoad;
            AfterAdd += this.OnAfterAdd;
            AfterUpdate += this.OnAfterUpdate;
            BeforeMenuClick += this.OnBeforeMenuClick;
            AfterItemClick += this.OnAfterItemClick;
            AfterDataLoad += this.OnAfterDataLoad;
        }

        private void OnBeforeMenuClick(object sender, ItemEventArgs e)
        {

            switch (e.ItemId)
            {
                case "1287":
                    this.AddOn.StatusBar.ErrorMessage = "The order cannot be duplicated.";
                    e.CancelEvent = true;
                    break;
            }
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public override void Dispose()
        {
            base.Dispose();
            FormResize -= this.OnFormResize;
            ButtonClick -= this.OnButtonClick;
            AfterFormLoad -= this.OnAfterFormLoad;
            AfterAdd -= this.OnAfterAdd;
            AfterUpdate -= this.OnAfterUpdate;
            BeforeMenuClick -= this.OnBeforeMenuClick;
            AfterItemClick -= this.OnAfterItemClick;
            AfterDataLoad -= this.OnAfterDataLoad;
        }

        private void OnAfterItemClick(object sender, ItemEventArgs e)
        {
            if ((e.ItemId == "38") && (e.ColId =="0"))
            {
                try
                {
                    this.DisableButtons();
                }
                catch
                {
                }
            }
        }

        private void OnAfterAdd(object sender, FormEventArgs e)
        {
            this.CreateSODirectory();
        }

        private void OnAfterDataLoad(object sender, FormEventArgs e)
        {
            this.ControlManager.Item("prodord").Enabled = false;
            this.ControlManager.Item("vieword").Enabled = false;
        }

        private void OnAfterUpdate(object sender, FormEventArgs e)
        {
            // We don't want to create directories if this is an update following a delete.
            if (!this._deleteInProgress)
            {
                this.CreateSODirectory();
            }

            this._deleteInProgress = false;
        }

        /// <summary>
        /// Called when [after form load].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        private void OnAfterFormLoad(object sender, FormEventArgs e)
        {
            try
            {
                SAPbouiCOM.ButtonCombo button = (SAPbouiCOM.ButtonCombo)this.SapForm.Items.Item("activity").Specific;
                if (button != null)
                {
                    button.ValidValues.Add("1", "Configure");
                    button.ValidValues.Add("2", "Commissions");
                    button.ValidValues.Add("3", "Duplicate");
                    button.ValidValues.Add("4", "Delete");
                    button.ValidValues.Add("5", "Copy Config.");
                    button.ValidValues.Add("6", "Open Excel");

                    button.Select(0, BoSearchKey.psk_Index);

                    this.SapForm.Items.Item("activity").DisplayDesc = true;
                }
            }
            catch { }
        }

        /// <summary>
        /// Called when [form resize].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        private void OnFormResize(object sender, FormEventArgs e)
        {
            // Set the buttons next to the cancel button
            UI.Adapters.Item activityBtn = this.ControlManager.Item("activity");
            var cancelBtn = new UI.Adapters.Button(this.SapForm, "2");

            activityBtn.MoveOnRightOf(cancelBtn, 5);

            UI.Adapters.Item approvalBtn = this.ControlManager.Item("approval");
            approvalBtn.MoveOnRightOf(activityBtn, 5);

            UI.Adapters.Item viewProdOrderBtn = this.ControlManager.Item("vieword");
            viewProdOrderBtn.MoveOnRightOf(approvalBtn, 5);

            UI.Adapters.Item createProdOrderBtn = this.ControlManager.Item("prodord");
            createProdOrderBtn.MoveOnRightOf(viewProdOrderBtn, 5);

            UI.Adapters.Item viewFolderBtn = this.ControlManager.Item("folder");
            viewFolderBtn.MoveOnRightOf(createProdOrderBtn, 5);
        }

        /// <summary>
        /// Raises the AfterItemClick event.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        protected void OnButtonClick(object sender, ItemEventArgs e)
        {
            this.SapForm.Items.Item(e.ItemId).Enabled = false;

            try
            {

                switch (e.ItemId)
                {
                    case "activity":
                        if (this.ValidateSelection())
                        {
                            this.SelectAction();
                        }
                        break;
                    case "folder":
                        if (this.ValidateSelection())
                        {
                            this.ViewDrawings();
                        }
                        break;
                    case "prodord":
                        if (this.ValidateSelection())
                        {
                            this.CreateProductionOrder();
                        }
                        break;
                    case "vieword":
                        if (this.ValidateSelection())
                        {
                            this.ShowProductionOrder();
                        }
                        break;
                    case "approval":
                        if (this.ValidateSelection())
                        {
                            this.OpenApproval();
                        }
                        break;
                }
            }
            finally
            {
                this.SapForm.Items.Item(e.ItemId).Enabled = true;
            }
        }

        private bool ValidateSelection()
        {
            try
            {
                var matrix = this.ControlManager.Matrix("38");
                if (matrix.SelectedRow() > -1)
                {
                    string rowType = matrix[matrix.SelectedRow(), 1];
                    if (rowType != string.Empty && rowType.ToUpper() != "R")
                    {
                        this.AddOn.StatusBar.ErrorMessage = "This action is not allowed on this row type.";
                        return false;
                    }
                }

                try
                {
                    var data = this.DataSource.SystemTable("RDR1");
                    int lineNum = this.GetSelectedLineNum();
                    int lineIndex = this.GetSelectedLine(int.Parse(this.PrimaryKey), lineNum);
                    string itemCode = data.GetValue("ItemCode", lineIndex);
                    if (string.IsNullOrEmpty(itemCode))
                    {
                        throw new Exception("Invalid Item Code");
                    }
                }
                catch
                {
                    this.AddOn.StatusBar.ErrorMessage = "The current line does not have a valid item selected.";
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
                return false;
            }
        }

        private void CreateSODirectory()
        {
            var data = this.DataSource.SystemTable("RDR1");
            object val = data.Fields.Item("DocEntry");

            int pKey;
            string primaryKey = this.PrimaryKey;
            if (int.TryParse(primaryKey, out pKey))
            {
                FortressSalesOrderAdapter.GetOrCreateSalesOrderDirectory(this.AddOn.Company, pKey);

                using (var soAdapter = new FortressSalesOrderAdapter(this.AddOn.Company, pKey))
                {
                    soAdapter.CreateAllSalesOrderLineDirectories();
                }
            }
            else
            {
                this.AddOn.StatusBar.ErrorMessage = "Could not determine sales order number";
            }
        }

        /// <summary>
        /// Selects the action.
        /// </summary>
        private void SelectAction()
        {
            try
            {
                SAPbouiCOM.ButtonCombo button = (SAPbouiCOM.ButtonCombo)this.SapForm.Items.Item("activity").Specific;

                switch (button.Selected.Description)
                {
                    case "Commissions":
                        this.OpenCommission();
                        break;
                    case "Duplicate":
                        System.Windows.Forms.DialogResult res = this.AddOn.MessageBox("Are you sure you want to duplicate the selected row?", System.Windows.Forms.MessageBoxButtons.YesNo);
                        if (res == System.Windows.Forms.DialogResult.Yes)
                        {
                            this.Duplicate();
                        }
                        break;
                    case "Copy Config.":
                        this.CopyConfiguration();
                        break;
                    case "Open Excel":
                        string filename = this.AddOn.OpenFileDialog("Excel files (*.xlsx)|*.xlsx", "");
                        this.AddOn.MessageBox(string.Format("The selected file is {0}", filename));
                        break;
                    case "Delete":
                        System.Windows.Forms.DialogResult res2 = this.AddOn.MessageBox("Are you sure you want to delete the selected row?", System.Windows.Forms.MessageBoxButtons.YesNo);
                        if (res2 == System.Windows.Forms.DialogResult.Yes)
                        {
                            this.Delete();
                        }
                        break;
                    case "Configure":
                        if (!this.OpenConveyor())
                        {
                            if (!this.OpenMetalDetector())
                            {
                                this.AddOn.StatusBar.ErrorMessage = "This item is not configurable";
                            }
                        }
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        private int GetSelectedLineNum()
        {
            int selectedRow = -1;
            var matrix = this.ControlManager.Matrix("38");
            selectedRow = matrix.SelectedRow();

            if (selectedRow <= 0)
            {
                throw new FortressException("No order detail selected.");
            }

            if (this.DataSource.SystemTable("RDR1").Size <= (selectedRow - 1)) return -1;

            return Int32.Parse(matrix[matrix.SelectedRow(), "110"]);
        }

        public void ViewDrawings()
        {
            int docEntry = int.Parse(this.PrimaryKey);

            if (docEntry <= 0)
            {
                throw new FortressException(string.Format("Can't load sales order {0} for production order creation. Has the sales order been saved?",
                                                          docEntry));
            }
            int lineNum = this.GetSelectedLineNum();
            int lineIndex = this.GetSelectedLine(docEntry, lineNum);

            DirectoryInfo directory;
            using (var soAdapter = new FortressSalesOrderAdapter(this.AddOn.Company, docEntry))
            {
                soAdapter.Lines.SetCurrentLine(lineIndex);
                directory = soAdapter.GetOrCreateSalesOrderLineDirectory(lineNum, soAdapter.Lines.Current.ItemCode);
                
            }

            if (directory != null && directory.Exists)
            {
                this.AddOn.OpenFolder(directory.FullName);
            }
            else
            {
                this.AddOn.StatusBar.ErrorMessage = "Cannot find sales order line directory [" + directory.FullName + "]";
                return;
            }
        }

        /// <summary>
        /// Creates the production order for the selected line.
        /// </summary>
        private void DisableButtons()
        {
            bool enableView = false;
            bool enableCreate = false;

            int docEntry = int.Parse(this.PrimaryKey);
            if (docEntry > 0)
            {
                using (var adapter = new FortressSalesOrderAdapter(this.AddOn.Company, docEntry))
                {
                    int currentLine = this.GetSelectedLineNum();
                    if (currentLine >= 0)
                    {
                        enableView = adapter.HasProductionOrder(currentLine);
                        enableCreate = !enableView;
                    }
                }
            }

            this.ControlManager.Item("prodord").Enabled = enableCreate;
            this.ControlManager.Item("vieword").Enabled = enableView;
        
        }

        /// <summary>
        /// Creates the production order for the selected line.
        /// </summary>
        public void CreateProductionOrder()
        {
            int docEntry = int.Parse(this.PrimaryKey);

            if (docEntry <= 0)
            {
                throw new FortressException(string.Format("Can't load sales order {0} for production order creation. Has the sales order been saved?",docEntry));
            }

            // Load the sales order
            using (var adapter = new FortressSalesOrderAdapter(this.AddOn.Company, docEntry))
            {
                int currentLine = this.GetSelectedLineNum();
                if (currentLine < 0)
                {
                    this.AddOn.StatusBar.ErrorMessage = "No order detail selected.";
                    return;
                }

                // Check if the line has been configured
                if (!adapter.IsLineConfigured(currentLine))
                {
                    this.AddOn.MessageBox("This line has not yet been configured. Please configure line before creating Production Order.");
                    return;
                }

                if (!adapter.HasProductionOrder(currentLine))
                {
                    try
                    {
                        // Create new producton orders
                        adapter.CreateProductionOrder(currentLine);
                    }
                    finally
                    {
                        // Refresh the screen
                        this.Refresh();
                        this.DisableButtons();
                    }
                }
            }

        }

        private int GetSelectedLine(int docEntry, int lineNum)
        {

            using (var rs = new RecordsetAdapter(this.AddOn.Company, string.Format("select VisOrder from RDR1 where DocEntry = {0} and LineNum = {1}", docEntry, lineNum)))
            {
                if (!rs.EoF)
                {
                    return int.Parse(rs.FieldValue("VisOrder").ToString());
                }
            }

            return -1;
        }

        /// <summary>
        /// Gets the selected line.
        /// </summary>
        /// <returns></returns>
        private int GetSelectedLine()
        {
            // Find selected line
            var rows = this.ControlManager.Matrix("38");
            if (rows.SelectedRow() == -1 && FortressConfigurationHelper.DoAutoSelectFirstRow(this.AddOn.Company))
            {
                rows.SelectFirstRow();
            }

            int selectedRow = rows.SelectedRow();

            return selectedRow;
        }

        /// <summary>
        /// Shows the production order.
        /// </summary>
        private void ShowProductionOrder()
        {
            int docEntry = int.Parse(this.PrimaryKey);
            if (docEntry > 0)
            {
                try
                {
                    using (var adapter = new FortressSalesOrderAdapter(this.AddOn.Company, docEntry))
                    {
                        int currentLine = this.GetSelectedLineNum();
                        if (currentLine >= 0)
                        {
                            IList<int> pos = adapter.GetProductionOrdersForLine(currentLine);
                            if (pos.Count == 0)
                            {
                                this.AddOn.StatusBar.ErrorMessage = "No production orders found for this line.";
                            }
                            else if (pos.Count == 1)
                            {
                                this.OpenDocument(BoLinkedObject.lf_ProductionOrder, pos[0]);
                            }
                            else
                            {
                                // Display screen
                                this.OpenSelectProductionOrders(currentLine);
                            }
                        }
                    }
                }
                finally
                {
                    this.DisableButtons();
                }
            }
        }

        private void OpenSelectProductionOrders(int lineNumber)
        {
            try
            {
                using (var adapter = new FortressSalesOrderAdapter(this.AddOn.Company, int.Parse(this.PrimaryKey)))
                {
                    foreach (var form in this.AddOn.Forms)
                    {
                        if (form.Value is SelectProductionOrders)
                        {
                            if ((((SelectProductionOrders)form.Value).OrderId == adapter.Document.DocNum.ToString())
                                && (((SelectProductionOrders)form.Value).OrderLine == lineNumber.ToString()))
                            {
                                ((SelectProductionOrders)form.Value).LoadProductionOrders();
                                ((SelectProductionOrders)form.Value).SapForm.Select();

                                return;
                            }
                        }
                    }

                    var approvalForm = this.AddOn.FormFactory.Create<SelectProductionOrders>("SelectProductionOrders.srf", "SPO");

                    var arguments = new Dictionary<string, object>
                                    {
                                        {"OrderId", adapter.Document.DocNum},
                                        {"OrderLine", lineNumber}
                                    };

                    approvalForm.Arguments = arguments;
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }


        ///// <summary>
        ///// Raises the <see cref="UI.Windows.Form.OnAfterDataLoad"/> event.
        ///// </summary>
        ///// <param name="sender">The sender.</param>
        ///// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        //private void OnAfterComboSelect(object sender, ItemEventArgs e)
        //{
        //    //this.SapForm.BusinessObject.Type;
        //    if (e.ItemId == "activCmb")
        //    {
        //        this.ControlManager.Button("activity").Value = this.ControlManager.ComboBox("activCmb").Value;
        //    }

        //}

        /// <summary>
        /// Opens the approval.
        /// </summary>
        private void OpenApproval()
        {
            try
            {
                var matrix = this.ControlManager.Matrix("38");

                if (this.GetSelectedLine() <= 0)
                {
                    this.AddOn.StatusBar.ErrorMessage = "No order detail selected.";
                    return;
                }

                var data = this.DataSource.SystemTable("RDR1");
                string rowNumber = matrix.SelectedRow().ToString();
                int lineNum = this.GetSelectedLineNum();
                int lineIndex = this.GetSelectedLine(int.Parse(this.PrimaryKey), lineNum);
                string approvalReg = data.GetValue("U_XX_ApprovalReq", lineIndex);

                if (approvalReg.Trim() == "No")
                {
                    this.AddOn.StatusBar.ErrorMessage = "This functionality is not avalable if approval is not required.";
                    return;
                }

                bool recordExists;
                using (var rs = new RecordsetAdapter(this.AddOn.Company, string.Format("select 1 from [@XX_APPROV] WHERE U_XX_OrderNo = '{0}' AND U_XX_OrdrLnNo = {1}", this.PrimaryKey, lineNum)))
                {
                    recordExists = !rs.EoF;
                }

                foreach (var form in this.AddOn.Forms)
                {
                    if (form.Value is Approval)
                    {
                        if ((((FortressForm)form.Value).OrderId == this.PrimaryKey)
                            && (((FortressForm)form.Value).OrderLine == lineNum.ToString()))
                        {
                            ((FortressForm)form.Value).SapForm.Select();

                            return;
                        }
                    }
                }

                var approvalForm = this.AddOn.FormFactory.Create<Approval>("Approval.srf", "AP");

                var arguments = new Dictionary<string, object>
                                    {
                                        {"RecordExists", recordExists},
                                        {"OrderId", this.PrimaryKey},
                                        {"DocumentNumber", this.DocumentNumber},
                                        {"OrderLine", lineNum},
                                        {"RowNumber", rowNumber}
                                    };

                approvalForm.Arguments = arguments;
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Opens the commission.
        /// </summary>
        private void OpenCommission()
        {
            try
            {
                var matrix = this.ControlManager.Matrix("38");

                if (this.GetSelectedLine() <= 0)
                {
                    this.AddOn.StatusBar.ErrorMessage = "No order detail selected.";
                    return;
                }

                var data = this.DataSource.SystemTable("RDR1");

                string lineNum = matrix[matrix.SelectedRow(), "110"];
                string rowNumber = matrix.SelectedRow().ToString();

                bool recordExists;
                using (var rs = new RecordsetAdapter(this.AddOn.Company, string.Format("select 1 from [@XX_COMMISS] WHERE U_XX_OrderNo = '{0}' AND U_XX_OrdrLnNo = {1}", this.PrimaryKey, lineNum)))
                {
                    recordExists = !rs.EoF;
                }

                foreach (var form in this.AddOn.Forms)
                {
                    if (form.Value is Commission)
                    {
                        if ((((FortressForm)form.Value).OrderId == this.PrimaryKey)
                            && (((FortressForm)form.Value).OrderLine == lineNum))
                        {
                            ((FortressForm)form.Value).SapForm.Select();

                            return;
                        }
                    }
                }

                var approvalForm = this.AddOn.FormFactory.Create<Commission>("Commission.srf", "CM");

                var arguments = new Dictionary<string, object>
                                    {
                                        {"RecordExists", recordExists},
                                        {"OrderId", this.PrimaryKey},
                                        {"DocumentNumber", this.DocumentNumber},
                                        {"OrderLine", lineNum},
                                        {"RowNumber", rowNumber}
                                    };

                approvalForm.Arguments = arguments;
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Determines whether this instance is a conveyor.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance is conveyor; otherwise, <c>false</c>.
        /// </returns>
        private bool IsConveyor()
        {
            var matrix = this.ControlManager.Matrix("38");
            var data = this.DataSource.SystemTable("RDR1");
            int lineNum = this.GetSelectedLineNum();
            int lineIndex = this.GetSelectedLine(int.Parse(this.PrimaryKey), lineNum);
            string itemCode = data.GetValue("ItemCode", lineIndex);

            string prefix;
            ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "ConvPref", out prefix);

            if (!itemCode.StartsWith(prefix))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Determines whether this instance is a conveyor.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance is conveyor; otherwise, <c>false</c>.
        /// </returns>
        private bool IsMetalDetector()
        {
            var matrix = this.ControlManager.Matrix("38");
            var data = this.DataSource.SystemTable("RDR1");
            int lineNum = this.GetSelectedLineNum();
            int lineIndex = this.GetSelectedLine(int.Parse(this.PrimaryKey), lineNum);

            string itemCode = data.GetValue("ItemCode", lineIndex);

            string prefix;
            ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "DetPref", out prefix);

            if (!itemCode.StartsWith(prefix))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Opens the conveyor.
        /// </summary>
        private bool OpenConveyor()
        {
            try
            {
                var matrix = this.ControlManager.Matrix("38");

                if (this.GetSelectedLine() <= 0)
                {
                    this.AddOn.StatusBar.ErrorMessage = "No order detail selected.";
                    return false;
                }

                var data = this.DataSource.SystemTable("RDR1");

                string lineNum = this.GetSelectedLineNum().ToString();
                string quantity = matrix[matrix.SelectedRow(), "9"];
                string itemCode = matrix[matrix.SelectedRow(), "1"];
                string rowNumber = matrix.SelectedRow().ToString();

                string prefix;
                ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "ConvPref", out prefix);

                if (!itemCode.StartsWith(prefix))
                {
                    return false;
                }

                bool recordExists;
                using (var rs = new RecordsetAdapter(this.AddOn.Company, string.Format("select 1 from [@XX_CONVBLT] WHERE U_XX_OrderNo = '{0}' AND U_XX_OrdrLnNo = {1}", this.PrimaryKey, lineNum)))
                {
                    recordExists = !rs.EoF;
                }

                foreach (var form in this.AddOn.Forms)
                {
                    if (form.Value is ConveyorConfigurator)
                    {
                        if ((((FortressForm)form.Value).OrderId == this.PrimaryKey)
                            && (((FortressForm)form.Value).OrderLine == lineNum))
                        {
                            ((FortressForm)form.Value).SapForm.Select();

                            return true;
                        }
                    }
                }

                var configForm = this.AddOn.FormFactory.Create<ConveyorConfigurator>("ConveyorConfigurator.srf", "CC");

                var arguments = new Dictionary<string, object>
                                    {
                                        {"RecordExists", recordExists},
                                        {"OrderId", this.PrimaryKey},
                                        {"DocumentNumber", this.DocumentNumber},
                                        {"OrderLine", lineNum},
                                        {"Quantity", quantity},
                                        {"RowNumber", rowNumber},
                                        {"ItemCode", itemCode}
                                    };

                configForm.Arguments = arguments;

                return true;
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
                return false;
            }
        }

        /// <summary>
        /// Opens the conveyor.
        /// </summary>
        private bool OpenMetalDetector()
        {
            try
            {
                var matrix = this.ControlManager.Matrix("38");

                if (this.GetSelectedLine() <= 0)
                {
                    this.AddOn.StatusBar.ErrorMessage = "No order detail selected.";
                    return false;
                }

                var data = this.DataSource.SystemTable("RDR1");

                string lineNum = this.GetSelectedLineNum().ToString();
                string quantity = matrix[matrix.SelectedRow(), "9"];
                string itemCode = matrix[matrix.SelectedRow(), "1"];
                string rowNumber = matrix.SelectedRow().ToString();

                //// Configuration can only happen if qty = 1
                //decimal qty;
                //if (decimal.TryParse(quantity, out qty))
                //{
                //    if (qty != 1.0m)
                //    {
                //        this.AddOn.StatusBar.ErrorMessage = "Line Quantity must be 1 to configure.";
                //        return false;
                //    }
                //}
                //else
                //{
                //    this.AddOn.StatusBar.ErrorMessage = "Line Quantity must be 1 to configure.";
                //    return false;
                //}

                string prefix;
                ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "DetPref", out prefix);

                if (!itemCode.StartsWith(prefix))
                {
                    return false;
                }

                bool recordExists;
                using (var rs = new RecordsetAdapter(this.AddOn.Company, string.Format("select 1 from [@XX_METDET] WHERE U_XX_OrderNo = '{0}' AND U_XX_OrdrLnNo = {1}", this.PrimaryKey, lineNum)))
                {
                    recordExists = !rs.EoF;
                }

                foreach (var form in this.AddOn.Forms)
                {
                    if (form.Value is ProductConfigurator)
                    {
                        if ((((FortressForm)form.Value).OrderId == this.PrimaryKey) && (((FortressForm)form.Value).OrderLine == lineNum))
                        {
                            ((FortressForm)form.Value).SapForm.Select();

                            return true;
                        }
                    }
                }

                var configForm = this.AddOn.FormFactory.Create<ProductConfigurator>("ProductConfigurator.srf", "PC");

                var arguments = new Dictionary<string, object>
                                    {
                                        {"RecordExists", recordExists},
                                        {"OrderId", this.PrimaryKey},
                                        {"DocumentNumber", this.DocumentNumber},
                                        {"OrderLine", lineNum},
                                        {"Quantity", quantity},
                                        {"RowNumber", rowNumber},
                                        {"ItemCode", itemCode}
                                    };

                configForm.Arguments = arguments;

                return true;
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
                return false;
            }
        }

        public bool CopyConfiguration()
        {
            try
            {
                var matrix = this.ControlManager.Matrix("38");

                if (this.GetSelectedLine() <= 0)
                {
                    this.AddOn.StatusBar.ErrorMessage = "No order detail selected.";
                    return true;
                }

                bool isConveyor = this.IsConveyor();
                bool isMetalDetector = this.IsMetalDetector();

                if (isConveyor || isMetalDetector)
                {
                    var data = this.DataSource.SystemTable("RDR1");

                    string rowNumber = matrix.SelectedRow().ToString();
                    int lineNum = this.GetSelectedLineNum();
                    int lineIndex = this.GetSelectedLine(int.Parse(this.PrimaryKey), lineNum);
                    string itemCode = data.GetValue("ItemCode", lineIndex);

                    // Do not allow copying from is already a configuration exists
                    string query = string.Empty;
                    if (isConveyor)
                    {
                        query = string.Format("SELECT 1 FROM [@XX_CONVBLT] WITH (NOLOCK) WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1} ", this.PrimaryKey, lineNum);
                    }
                    else
                    {
                        query = string.Format("SELECT 1 FROM [@XX_METDET] WITH (NOLOCK) WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1} ", this.PrimaryKey, lineNum);
                    }

                    using (var rs = new RecordsetAdapter(this.AddOn.Company, query))
                    {
                        if (!rs.EoF)
                        {
                            this.AddOn.StatusBar.ErrorMessage = "A configuration for the selected line already exists.";
                            return true;
                        }
                    }

                    foreach (var form in this.AddOn.Forms)
                    {
                        if (form.Value is CopyConfigurator)
                        {
                            if ((((CopyConfigurator)form.Value).OrderId == this.PrimaryKey)
                                && (((CopyConfigurator)form.Value).OrderLine == lineNum.ToString()))
                            {
                                ((CopyConfigurator)form.Value).SapForm.Select();

                                return true;
                            }
                        }
                    }

                    var copyConfig = this.AddOn.FormFactory.Create<CopyConfigurator>("CopyConfigurator.srf", "CC");

                    var arguments = new Dictionary<string, object>
                                    {
                                        {"OrderId", this.PrimaryKey},
                                        {"OrderLine", lineNum},
                                        {"RowNumber", rowNumber},
                                        {"ItemCode", itemCode}
                                    };

                    copyConfig.Arguments = arguments;
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }

            return true;
        }

        /// <summary>
        /// Duplicates the selected row (Configuration & Commision)
        /// </summary>
        private bool Duplicate()
        {
            if (this.IsConveyor())
            {
                if (this.DuplicateConveyorBelt())
                {
                    return true;
                }
                return false;
            }
            else if (this.IsMetalDetector())
            {
                if (this.DuplicateMetalDetector())
                {
                    return true;
                }
                return false;
            }
            else
            {
                return this.DuplicateNonConfiguredItem();
            }
        }

        private bool Delete()
        {
            this._deleteInProgress = true;

            try
            {
                int currentLineNum = this.GetSelectedLineNum();

                int selectedRow = -1;
                var matrix = this.ControlManager.Matrix("38");
                selectedRow = matrix.SelectedRow();

                if (selectedRow <= 0)
                {
                    this.AddOn.StatusBar.ErrorMessage = "No order detail selected.";
                    return true;
                }



                using (var soAdapter = new FortressSalesOrderAdapter(this.AddOn.Company, int.Parse(this.PrimaryKey)))
                {
                    soAdapter.SetLine(currentLineNum);
                    DirectoryInfo dInfo = soAdapter.GetSalesOrderLineDirectory(currentLineNum, soAdapter.Lines.Current.ItemCode);
                    if (dInfo != null)
                    {
                        dInfo.MoveTo(dInfo.Parent.FullName + "\\" + soAdapter.Lines.Current.ItemCode + "-" + DateTime.Now.ToString("yyyyMMdd-hhmmss"));
                    }
                }

                //Select the cell in edit mode to force the
                // Duplicate Rom Menu to be activated
                matrix[selectedRow, 3] = matrix[selectedRow, 3];

                // Delete children records
                const string Query = "DELETE FROM [@XX_CONVBLT] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}; DELETE FROM [@XX_APPROV] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}; DELETE FROM [@XX_COMMISS]  WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}; DELETE FROM [@XX_METDET] WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}";
                using (var rs = new RecordsetAdapter(this.AddOn.Company))
                {
                    rs.Execute(string.Format(Query, this.PrimaryKey, currentLineNum));
                }

                // Record is valid for duplication
                // Run the delete row menu click
                this.AddOn.Application.SendKeys("(^K)");

                // Save order
                this.ControlManager.Button("1").Focus();
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex); 
            }
            

            return true;
        }

        private bool DuplicateCommision(string originalLineNum, string newLine)
        {
            try
            {
                var matrix = this.ControlManager.Matrix("38");

                var data = this.DataSource.SystemTable("RDR1");

                var loadData = new System.Text.StringBuilder();
                loadData.AppendLine("<DataLoader UDOName='Commission' UDOType='MasterData'>");

                // Get all the data from the current commision:
                string query = string.Format(@"SELECT DocEntry, Code, U_XX_Type as ComType, Name, U_XX_Rate as Rate,  CONVERT(varchar(10), CreateDate, 121) as CreateDate, LEFT(CreateTime, len(CreateTime)-2) + ':' + RIGHT(CreateTime, 2) as CreateTime, CONVERT(varchar(10), UpdateDate, 121) as UpdateDate, LEFT(UpdateTime, len(UpdateTime)-2) + ':' + RIGHT(UpdateTime, 2) as UpdateTime, U_XX_Approved  as Approved FROM [@XX_COMMISS] WITH (NOLOCK) WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}", this.PrimaryKey, originalLineNum);

                using (var recordset = new RecordsetAdapter(this.AddOn.Company, query))
                {
                    // If there are no commission lines, return true.
                    if (recordset.EoF)
                    {
                        return true;
                    }

                    int counter = 0;
                    while (!recordset.EoF)
                    {
                        string nextCode = new TableFieldAdapter(this.AddOn.Company, "@XX_COMMISS", "Code").NextCode("CMM", string.Empty, counter);

                        loadData.AppendLine("<Record>");
                        loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(nextCode));
                        loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(recordset.FieldValue("Name").ToString()));
                        loadData.AppendFormat("<Field Name='U_XX_Rate' Type='System.Double' Value='{0}' />", HttpUtility.HtmlEncode(recordset.FieldValue("Rate").ToString()));
                        loadData.AppendFormat("<Field Name='U_XX_Type' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(recordset.FieldValue("ComType").ToString()));
                        loadData.AppendFormat("<Field Name='U_XX_OrderNo' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(this.PrimaryKey));
                        loadData.AppendFormat("<Field Name='U_XX_OrdrLnNo' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(newLine));
                        loadData.AppendFormat("<Field Name='U_XX_Approved' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(recordset.FieldValue("Approved").ToString()));
                        loadData.AppendLine("</Record>");

                        counter++;
                        recordset.MoveNext();
                    }
                }

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
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
                return false;
            }

            return true;
        }

        
        /// <summary>
        /// Duplicates the metal detector.
        /// </summary>
        /// <returns></returns>
        private bool DuplicateNonConfiguredItem()
        {
            try
            {
                int selectedRow = -1;
                var matrix = this.ControlManager.Matrix("38");
                selectedRow = matrix.SelectedRow();

                if (selectedRow <= 0)
                {
                    this.AddOn.StatusBar.ErrorMessage = "No order detail selected.";
                    return true;
                }

                //Select the cell in edit mode to force the
                // Duplicate Rom Menu to be activated
                matrix[selectedRow, 3] = matrix[selectedRow, 3];

                var data = this.DataSource.SystemTable("RDR1");
                int lineNum = this.GetSelectedLineNum();
                int lineIndex = this.GetSelectedLine(int.Parse(this.PrimaryKey), lineNum);

                string origOrderNo = data.GetValue("DocEntry", lineIndex);
                string origOrderLineNo = data.GetValue("LineNum", lineIndex);
                string quantity = data.GetValue("Quantity", lineIndex);
                string itemCode = data.GetValue("ItemCode", lineIndex);
               
                // Record is valid for duplication
                // Run the duplicate row menu click
                if (!this.AddOn.Application.Menus.Item("1294").Enabled)
                {
                    this.AddOn.Application.SendKeys("(^M)");
                }

                this.AddOn.Application.SendKeys("(^M)");
                this.ControlManager.Button("1").Focus();

                string newLine = string.Empty;
                // Get the lates line number
                using (var rs = new RecordsetAdapter(this.AddOn.Company, string.Format("SELECT Max(LineNum) as NewLineNum FROM RDR1 WHERE DocEntry = {0}", origOrderNo)))
                {
                    if (rs.EoF)
                    {
                        this.AddOn.StatusBar.ErrorMessage = string.Format("The order {0} is not found.", origOrderNo);
                        return false;
                    }

                    newLine = rs.FieldValue("NewLineNum").ToString();
                }

                using (var soAdapter = new FortressSalesOrderAdapter(this.AddOn.Company, int.Parse(this.PrimaryKey)))
                {
                    soAdapter.DuplicateCommission(int.Parse(origOrderLineNo), int.Parse(newLine));
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Duplicates the metal detector.
        /// </summary>
        /// <returns></returns>
        private bool DuplicateMetalDetector()
        {

            try
            {
                int selectedRow = -1;
                var matrix = this.ControlManager.Matrix("38");
                selectedRow = matrix.SelectedRow();

                if (selectedRow <= 0)
                {
                    this.AddOn.StatusBar.ErrorMessage = "No order detail selected.";
                    return true;
                }

                //Select the cell in edit mode to force the
                // Duplicate Rom Menu to be activated
                matrix[selectedRow, 3] = matrix[selectedRow, 3];

                var data = this.DataSource.SystemTable("RDR1");
                int lineNum = this.GetSelectedLineNum();
                int lineIndex = this.GetSelectedLine(int.Parse(this.PrimaryKey), lineNum);

                string origOrderNo = data.GetValue("DocEntry", lineIndex);
                string origOrderLineNo = data.GetValue("LineNum", lineIndex);
                string quantity = data.GetValue("Quantity", lineIndex);
                string itemCode = data.GetValue("ItemCode", lineIndex);
                string prefix;
                ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "DetPref", out prefix);

                if (!itemCode.StartsWith(prefix))
                {
                    return false;
                }


                // Reset the production status after copying
                string prodStatus = data.GetValue("U_XX_ProdStatus", lineIndex);
                if (!string.IsNullOrEmpty(prodStatus))
                {
                    try
                    {
                        ((SAPbouiCOM.EditText)matrix.Columns.Item("U_XX_ProdStatus").Cells.Item(selectedRow).Specific).Value = string.Empty;
                    }
                    catch (Exception)
                    {
                        prodStatus = ((SAPbouiCOM.ComboBox)matrix.Columns.Item("U_XX_ProdStatus").Cells.Item(selectedRow).Specific).Selected.Value;
                        ((SAPbouiCOM.ComboBox)matrix.Columns.Item("U_XX_ProdStatus").Cells.Item(selectedRow).Specific).Select(null, BoSearchKey.psk_Index);
                    }
                }

                // Record is valid for duplication
                // Run the duplicate row menu click
                if (!this.AddOn.Application.Menus.Item("1294").Enabled)
                {
                    this.AddOn.Application.SendKeys("(^M)");
                }

                this.AddOn.Application.SendKeys("(^M)");

                // Save order
                if (!string.IsNullOrEmpty(prodStatus))
                {
                    try
                    {
                        ((SAPbouiCOM.EditText)matrix.Columns.Item("U_XX_ProdStatus").Cells.Item(selectedRow).Specific).Value = prodStatus;
                    }
                    catch (Exception)
                    {
                        ((SAPbouiCOM.ComboBox)matrix.Columns.Item("U_XX_ProdStatus").Cells.Item(selectedRow).Specific).Select(prodStatus, BoSearchKey.psk_ByValue);
                    }
                }

                this.ControlManager.Button("1").Focus();

                string newLine = string.Empty;
                // Get the lates line number
                using (var rs = new RecordsetAdapter(this.AddOn.Company, string.Format("SELECT Max(LineNum) as NewLineNum FROM RDR1 WHERE DocEntry = {0}", origOrderNo)))
                {
                    if (rs.EoF)
                    {
                        this.AddOn.StatusBar.ErrorMessage = string.Format("The order {0} is not found.", origOrderNo);
                        return false;
                    }

                    newLine = rs.FieldValue("NewLineNum").ToString();
                }

                using (var soAdapter = new FortressSalesOrderAdapter(this.AddOn.Company, int.Parse(this.PrimaryKey)))
                {
                    soAdapter.DuplicateConfiguration(int.Parse(origOrderLineNo), int.Parse(newLine));
                }
                return true;
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
                return false;
            }
        }

        /// <summary>
        /// Duplicates the conveyor belt.
        /// </summary>
        /// <returns></returns>
        private bool DuplicateConveyorBelt()
        {

            try
            {
                int selectedRow = -1;
                var matrix = this.ControlManager.Matrix("38");
                selectedRow = matrix.SelectedRow();

                if (selectedRow <= 0)
                {
                    this.AddOn.StatusBar.ErrorMessage = "No order detail selected.";
                    return true;
                }

                //Select the cell in edit mode to force the
                // Duplicate Rom Menu to be activated
                matrix[selectedRow, 3] = matrix[selectedRow, 3];

                var data = this.DataSource.SystemTable("RDR1");
                int lineNum = this.GetSelectedLineNum();
                int lineIndex = this.GetSelectedLine(int.Parse(this.PrimaryKey), lineNum);

                string origOrderNo = data.GetValue("DocEntry", lineIndex);
                string origOrderLineNo = data.GetValue("LineNum", lineIndex);
                string quantity = data.GetValue("Quantity", lineIndex);
                string itemCode = data.GetValue("ItemCode", lineIndex);

                string prefix;
                ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "ConvPref", out prefix);

                if (!itemCode.StartsWith(prefix))
                {
                    return false;
                }

                // Reset the production status after copying
                string prodStatus = data.GetValue("U_XX_ProdStatus", lineIndex);
                if (!string.IsNullOrEmpty(prodStatus))
                {
                    try
                    {
                        ((SAPbouiCOM.EditText)matrix.Columns.Item("U_XX_ProdStatus").Cells.Item(selectedRow).Specific).Value = string.Empty;
                    }
                    catch (Exception)
                    {
                        prodStatus = ((SAPbouiCOM.ComboBox)matrix.Columns.Item("U_XX_ProdStatus").Cells.Item(selectedRow).Specific).Selected.Value;
                        ((SAPbouiCOM.ComboBox)matrix.Columns.Item("U_XX_ProdStatus").Cells.Item(selectedRow).Specific).Select(null, BoSearchKey.psk_Index);
                    }
                }

                // Record is valid for duplication
                // Run the duplicate row menu click
                if (!this.AddOn.Application.Menus.Item("1294").Enabled)
                {
                    this.AddOn.Application.SendKeys("(^M)");
                }

                this.AddOn.Application.SendKeys("(^M)");

                if (!string.IsNullOrEmpty(prodStatus))
                {
                    try
                    {
                        ((SAPbouiCOM.EditText)matrix.Columns.Item("U_XX_ProdStatus").Cells.Item(selectedRow).Specific).Value = prodStatus;
                    }
                    catch (Exception)
                    {
                        ((SAPbouiCOM.ComboBox)matrix.Columns.Item("U_XX_ProdStatus").Cells.Item(selectedRow).Specific).Select(prodStatus, BoSearchKey.psk_ByValue);
                    }
                }

                // Save order
                this.ControlManager.Button("1").Focus();

                string newLine = string.Empty;
                // Get the lates line number
                using (var rs = new RecordsetAdapter(this.AddOn.Company, string.Format("SELECT Max(LineNum) as NewLineNum FROM RDR1 WHERE DocEntry = {0}", origOrderNo)))
                {
                    if (rs.EoF)
                    {
                        this.AddOn.StatusBar.ErrorMessage = string.Format("The order {0} is not found.", origOrderNo);
                        return false;
                    }

                    newLine = rs.FieldValue("NewLineNum").ToString();
                }

                using (var soAdapter = new FortressSalesOrderAdapter(this.AddOn.Company, int.Parse(this.PrimaryKey)))
                {
                    soAdapter.DuplicateConfiguration(int.Parse(origOrderLineNo), int.Parse(newLine));
                }

                /*
                // Duplicate row
                using (var rs = new RecordsetAdapter(this.AddOn.Company, string.Format("SELECT * FROM [@XX_CONVBLT] WHERE U_XX_OrderNo = '{0}' and U_XX_OrdrLnNo = '{1}' and U_XX_Status = 'Initiated'", origOrderNo, origOrderLineNo)))
                {
                    if (rs.EoF)
                    {
                        this.AddOn.StatusBar.ErrorMessage = "The selected order line is not configured or not initialized.";
                        return false;
                    }

                    string newCode = origOrderNo + '-' + newLine;

                    var loadData = new System.Text.StringBuilder();
                    loadData.AppendLine("<DataLoader UDOName='CBConfigurator' UDOType='MasterData'>");
                    loadData.AppendLine("<Record>");
                    loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(newCode));
                    loadData.AppendFormat("<Field Name='U_XX_OrderNo' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(origOrderNo));
                    loadData.AppendFormat("<Field Name='U_XX_OrdrLnNo' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(newLine));
                    loadData.AppendFormat("<Field Name='U_XX_Revis' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode("1"));
                    loadData.AppendFormat("<Field Name='U_XX_Status' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode("Initiated"));
                    loadData.AppendFormat("<Field Name='U_XX_OvLenght' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OvLenght").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltType' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltSpeed' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltSpeed").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltWd' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltWd").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltPM' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltPM").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltLn' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltLn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltDir' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltDir").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltHgIn' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltHgIn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltHgOut' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltHgOut").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_BltAdj' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_BltAdj").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_DetLoc' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_DetLoc").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_DetLocC' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_DetLocC").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_DetFrCtr' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_DetFrCtr").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_MntType' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_MntType").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_MntStyle' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_MntStyle").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_MntMat' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_MntMat").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutRReq' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutRReq").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutRStl' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutRStl").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutLn' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutLn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutLnU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutLnU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutWd' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutWd").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutWdU' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutWdU").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_OutHg' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_OutHg").ToString()));
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
                    loadData.AppendFormat("<Field Name='U_XX_FlGtLn' Type='System.Int32' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_FlGtLn").ToString()));
                    loadData.AppendFormat("<Field Name='U_XX_FlChute' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(rs.FieldValue("U_XX_FlChute").ToString()));

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
                        else
                        {
                            this.DuplicateCommision(origOrderLineNo, newLine);
                        }
                    }
                }
                */
                return true;
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
                return false;
            }
        }

        #endregion Method(s)
    }
}