//-----------------------------------------------------------------------
// <copyright file="Approval.cs" company="Fortress Technology Inc.">
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
    using B1C.SAP.DI.BusinessAdapters;
    using B1C.SAP.DI.Business;
    using Fortress.DI.Model.Exceptions;
    using System.Web;

    #endregion Usings

    public class Approval : Window.FortressForm
    {
        #region Constuctors

        /// <summary>
        /// Initializes a new instance of the <see cref="Approval"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        public Approval(UI.AddOn addOn, string xmlPath, string formType) : base(addOn, xmlPath, formType) { }

        #endregion Constuctors

        #region Methods

        /// <summary>
        /// Initializes this instance.
        /// </summary>
        public override void Initialize()
        {
            this.ButtonClick += this.OnButtonClick;
            this.AfterArgumentsSet += this.OnAfterArgumentsSet;
            this.FormResize += this.OnFormResize;
            this.AfterComboSelect += this.OnAfterComboSelect;
            this.BeforeComboSelect += this.OnBeforeComboSelect;
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
            this.AfterComboSelect -= this.OnAfterComboSelect;
            this.BeforeComboSelect -= this.OnBeforeComboSelect;
            base.Dispose();
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
                if (matrix[e.Row, "col_1"] != "<New>")
                {
                    e.CancelEvent = true;
                    e.CancelMessage = "User not autorized to perform this operation on an existing record.";
                }
            }
        }

        private void OnAfterComboSelect(object sender, ItemEventArgs e)
        {
            if (e.ItemId == "mtx_0")
            {
                this.ControlManager.Matrix("mtx_0")[e.Row, "col_5"] = "Y";
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
            matrix.Columns.Item(2).Width = 0;
            matrix.Columns.Item(4).Width = 80;
            matrix.Columns.Item(5).Width = 60;
            matrix.Columns.Item(6).Width = 0;
            matrix.FillWidthOnColumn(3);
        }

        /// <summary>
        /// Called when [after arguments set].
        /// </summary>
        /// <param name="sender">The sender.</param>
        private void OnAfterArgumentsSet(object sender)
        {
            this.ControlManager.EditText("soNumber").Value = this.OrderId;
            this.Field("soDocNoDS").Value = this.DocumentNumber;
            this.ControlManager.EditText("soLineNum").Value = this.OrderLine;
            this.ControlManager.EditText("soRowNum").Value = this.RowNumber;

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
        }

        private void DeleteRow()
        {
            var matrix = this.ControlManager.Matrix("mtx_0");

            if (matrix != null)
            {
                if (matrix.SelectedRow() != -1)
                {
                    string docEntry = matrix[matrix.SelectedRow(), "col_1"];

                    RecordsetAdapter.Execute(
                            this.AddOn.Company,
                            string.Format("DELETE FROM [@XX_APPROV] WHERE Code = '{0}'", docEntry));
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
        /// Called when [button click].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.ItemEventArgs"/> instance containing the event data.</param>
        private void OnButtonClick(object sender, ItemEventArgs e)
        {
            switch (e.ItemId)
            {
                case "3":
                    this.AddStatus();
                    break;
                case "1":
                    this.UpdateRecord();
                    break;
                case "delRow":
                    this.DeleteRow();
                    break;
            }
        }

        /// <summary>
        /// Adds the status.
        /// </summary>
        private void AddStatus()
        {
            try
            {
                var selectedSerial = this.DataSource.UserTable("StatusDT");

                if (selectedSerial != null)
                {
                    // Check the value of the last row
                    var lastCode = selectedSerial.GetValue("Code", selectedSerial.Rows.Count - 1);
                    if (!string.IsNullOrEmpty(lastCode.ToString()))
                    {
                        selectedSerial.Rows.Add(1);
                    }
                    selectedSerial.SetValue("DocEntry", selectedSerial.Rows.Count - 1, new TableFieldAdapter(this.AddOn.Company, "@XX_APPROV", "DocEntry").NextCode(string.Empty, string.Empty));
                    selectedSerial.SetValue("Code", selectedSerial.Rows.Count - 1, "<New>");
                    selectedSerial.SetValue("ApStatus", selectedSerial.Rows.Count - 1, "-");
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
                var approvalMatrix = this.ControlManager.Matrix("mtx_0");

                if (approvalMatrix != null)
                {
                    for (int rowIndex = 1; rowIndex <= approvalMatrix.RowCount; rowIndex++)
                    {
                        string status = approvalMatrix[rowIndex, "col_2"];
                        if (status == "-")
                        {
                            this.AddOn.MessageBox("Please select a valid status. Value '-' not allowed.");
                            //this.AddOn.StatusBar.ErrorMessage = "Please select a valid status. Value '-' not allowed.";
                            //this.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            return;
                        }
                    }

                    // Check the value of the last row
                    for (int rowIndex = 1; rowIndex <= approvalMatrix.RowCount; rowIndex++)
                    {
                        var lastCode = approvalMatrix[rowIndex, "col_1"];
                        if (lastCode == "<New>")
                        {
                            // Append New records
                            var loadData = new System.Text.StringBuilder();
                            loadData.AppendLine("<DataLoader UDOName='Approval' UDOType='MasterData'>");
                            loadData.AppendLine("<Record>");
                            loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", new TableFieldAdapter(this.AddOn.Company, "@XX_APPROV", "Code").NextCode("APP", string.Empty));
                            loadData.AppendFormat("<Field Name='U_XX_Status' Type='System.String' Value='{0}' />", HttpUtility.HtmlEncode(approvalMatrix[rowIndex, "col_2"]));
                            loadData.AppendFormat("<Field Name='U_XX_OrderNo' Type='System.Int32' Value='{0}' />", this.OrderId);
                            loadData.AppendFormat("<Field Name='U_XX_OrdrLnNo' Type='System.Int32' Value='{0}' />", this.OrderLine);
                            loadData.AppendLine("</Record>");
                            loadData.AppendLine("</DataLoader>");

                            DataLoader loader = DataLoader.InstanceFromXMLString(loadData.ToString());

                            if (loader != null)
                            {
                                loader.Process(this.AddOn.Company);
                            }
                        }
                        else
                        {
                            // Check if record is updated
                            if (approvalMatrix[rowIndex, "col_5"] == "Y")
                            {
                                var statuses = this.AddOn.Company.UserTables.Item("XX_APPROV");
                                statuses.GetByKey(lastCode);
                                statuses.UserFields.Fields.Item("U_XX_Status").Value = approvalMatrix[rowIndex, "col_2"];
                                statuses.Update();
                            }
                        }
                    }
                }


                this.LoadMatrix();

                this.EnabledButton("delRow", true);
                this.EnabledButton("3", true);
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }


        private void LoadMatrix()
        {
            var statusTable = this.DataSource.UserTable("StatusDT");
            statusTable.Rows.Clear();

            string query = string.Format(@"SELECT DocEntry, Code, U_XX_Status as ApStatus, CONVERT(varchar(10), CreateDate, 121) as CreateDate, LEFT(CreateTime, len(CreateTime)-2) + ':' + RIGHT(CreateTime, 2) as CreateTime  FROM [@XX_APPROV] WITH (NOLOCK) WHERE U_XX_OrderNo = {0} AND U_XX_OrdrLnNo = {1}", this.OrderId, this.OrderLine);

            statusTable.ExecuteQuery(query);

            this.ControlManager.Matrix("mtx_0").BaseObject.LoadFromDataSource();

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

        #endregion Methods
    }
}