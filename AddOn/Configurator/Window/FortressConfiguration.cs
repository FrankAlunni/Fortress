//-----------------------------------------------------------------------
// <copyright file="FortressConfiguration.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
// <email>frank.alunni@b1computing.com</email>
//-----------------------------------------------------------------------

using System;
using System.Collections;
using B1C.SAP.UI.Model.EventArguments;
using B1C.Utility.Helpers;
using Fortress.DI.BusinessAdapters.Sales;
using B1C.SAP.DI.BusinessAdapters;
using Fortress.DI.BusinessAdapters;
using System.Windows.Forms;
namespace B1C.SAP.Addons.Configurator.Window
{
    #region Using Directive(s)

    #endregion Using Directive(s)

    public abstract class FortressConfiguration : FortressForm
    {
        public enum ConfigurationType
        {
            MetalDetector,
            ConveyorBelt
        }

        #region Properties
        public abstract ConfigurationType FormType { get; }

        /// <summary>
        /// Gets or sets the find code.
        /// </summary>
        /// <value>The find code.</value>
        protected string FindCode { get; set; }

        /// <summary>
        /// Gets the quantity.
        /// </summary>
        /// <value>The quantity.</value>
        protected double Quantity
        {
            get
            {
                object argument;
                this.Arguments.TryGetValue("Quantity", out argument);
                if (argument != null)
                {
                    return double.Parse(argument.ToString());
                }

                return 0;
            }
        }

        #endregion Properties

        #region Constuctors

        /// <summary>
        /// Initializes a new instance of the <see cref="FortressConfiguration"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        public FortressConfiguration(UI.AddOn addOn, string xmlPath, string formType) : base(addOn, xmlPath, formType) { }

        #endregion Constuctors

        #region Methods

        /// <summary>
        /// Initializes this instance.
        /// </summary>
        public override void Initialize()
        {
            AfterArgumentsSet += this.OnAfterArgumentsSet;
            BeforeUpdate += this.OnBeforeUpdate;
            AfterSave += this.OnAfterSave;
            base.Initialize();
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public override void Dispose()
        {
            AfterArgumentsSet -= this.OnAfterArgumentsSet;
            BeforeUpdate -= this.OnBeforeUpdate;
            AfterSave -= this.OnAfterSave;
            base.Dispose();
        }

        private void OnAfterSave(object sender, FormEventArgs e)
        {
            this.UpdateConfigurationComments();
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

            // Find the record
            if (!string.IsNullOrEmpty(this.FindCode))
            {
                this.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                this.ControlManager.EditText("code").Value = this.FindCode;
                this.ControlManager.Button("1").Focus();
            }

            this.ControlManager.EditText("quantity").Value = this.Quantity.ToString();
            this.ControlManager.EditText("soRowNum").Value = this.RowNumber.ToString();
            this.ControlManager.EditText("partNo").Value = this.ItemCode;

            
            
            using (var salesOrder = new FortressSalesOrderAdapter(this.AddOn.Company, int.Parse(this.OrderId)))
            {

                // If the SO is closed, show the configuration as read-only
                if (salesOrder.Document.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                {                    
                    try
                    {
                        // Disable all except buttons and controls
                        for (int controlIndex = 0; controlIndex < this.SapForm.Items.Count; controlIndex++)
                        {
                            if ((this.SapForm.Items.Item(controlIndex).Type != SAPbouiCOM.BoFormItemTypes.it_BUTTON) &&
                                (this.SapForm.Items.Item(controlIndex).Type != SAPbouiCOM.BoFormItemTypes.it_FOLDER))
                            {
                                this.SapForm.Items.Item(controlIndex).Enabled = false;
                            }
                        }
                    }
                    catch
                    {
                        // Ignore error in case the control has focus while disabling
                    }
                }

                if (salesOrder.HasProductionOrder(int.Parse(this.OrderLine)))
                {
                    this.SapForm.Title = this.SapForm.Title + " - [Production Order exists]";                    
                }
            }
        }

        /// <summary>
        /// Gets the original status.
        /// </summary>
        /// <returns></returns>
        private string GetOriginalStatus()
        {
            string originalStatus = "Initiated";
            string code = this.Field(this.TableName, "Code").Value;
            string query = string.Format("SELECT U_XX_Status FROM [{0}] WHERE Code = '{1}'", this.TableName, code);
            using (var recordset = new RecordsetAdapter(this.AddOn.Company, query))
            {
                if (!recordset.EoF)
                {
                    originalStatus = recordset.FieldValue("U_XX_Status").ToString();
                }
            }

            return originalStatus;
        }

        /// <summary>
        /// Called when [before update].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        private void OnBeforeUpdate(object sender, FormEventArgs e)
        {
            using (var salesOrder = new FortressSalesOrderAdapter(this.AddOn.Company, int.Parse(this.OrderId)))
            {
                if (salesOrder.HasProductionOrder(int.Parse(this.OrderLine)))
                {
                    DialogResult result = this.AddOn.MessageBox("A production order already exists. Are you sure you want to save changes to the configuration of this item?", System.Windows.Forms.MessageBoxButtons.YesNo);
                    if (result != DialogResult.Yes)
                    {
                        e.CancelEvent = true;
                        return;
                    }
                }
            }

            string newStatus = this.Field(this.TableName, "U_XX_Status").Value;
            string revision = this.Field(this.TableName, "U_XX_Revis").Value;

            string originalStatus = this.GetOriginalStatus();

            if (originalStatus == "Failed" || originalStatus == "Passed")
            {
                // Increase revision
                this.Field(this.TableName, "U_XX_Revis").Value = (int.Parse(revision) + 1).ToString();

                // Set status = 'Modified' - This should always happen unless the status has been manually modified
                if (newStatus == originalStatus)
                {
                    this.Field(this.TableName, "U_XX_Status").Value = "Modified";
                }
            }

            // Send out an alert
            if (int.Parse(revision) > 0)
            {
                FortressMessageAdapter.SendConfigurationUpdateMessage(this.AddOn.Company, int.Parse(this.OrderId), this.DocumentNumber,
                                                                      int.Parse(this.OrderLine), this.ItemCode, int.Parse(revision) + 1);
            }
        }


        /// <summary>
        /// Update the system notes by calling a stored procedure
        /// </summary>
        private void UpdateConfigurationComments()
        {
            if (this.FormType == ConfigurationType.MetalDetector)
            {
                // < 110413 - Dom
                try
                {

                    var sp = new B1C.Utility.Helpers.StoredProcHelper(this.AddOn.Connection.Database + "_FTI.dbo.FTI_ConfigUpdateDetectorNotes");
                    sp.AddParameter("@code", System.Data.DbType.String, this.ControlManager.EditText("code").Value);

                    RecordsetAdapter.Execute(this.AddOn.Company, sp.ToString());
                }
                catch (Exception ex)
                {
                    this.AddOn.StatusBar.ErrorMessage = "[Dom] Failed to run stored procedure [FTI_ConfigUpdateDetectorNotes] to update system notes. Error:" + ex.Message ;
                }

                // > 110413 - Dom

                /*  -- This is how Frank originally set it up, using a query.
                const string UpdateQuery = "UPDATE dbo.[@XX_METDET] SET U_XX_SNotes = CASE ISNULL(CAST(U_XX_SNotes as VARCHAR), '') WHEN '' THEN 'Model: ' + CAST(U_XX_Model as VARCHAR) ELSE CAST(U_XX_SNotes AS VARCHAR) END WHERE Code = '{0}'";

                using (var rs = new RecordsetAdapter(this.AddOn.Company))
                {
                    rs.Execute(string.Format(UpdateQuery, this.ControlManager.EditText("code").Value));
                }
                */
            }
            else
            {

                // < 110413 - Dom
                try
                {
                    var sp = new B1C.Utility.Helpers.StoredProcHelper(this.AddOn.Connection.Database + "_FTI.dbo.FTI_ConfigUpdateConveyorNotes");
                    sp.AddParameter("@code", System.Data.DbType.String, this.ControlManager.EditText("code").Value);

                    RecordsetAdapter.Execute(this.AddOn.Company, sp.ToString());

                }
                catch (Exception ex)
                {
                    this.AddOn.StatusBar.ErrorMessage = "[Dom] Failed to run stored procedure [FTI_ConfigUpdateConveyorNotes] tto update system notes. Error:" + ex.Message;
                }

                /*               
                const string UpdateQuery = "UPDATE dbo.[@XX_CONVBLT] SET U_XX_SNotes = CASE ISNULL(CAST(U_XX_SNotes as VARCHAR), '') WHEN '' THEN 'Size: ' + CAST(CAST(U_XX_OvLenght as decimal(9,2)) as varchar) +' X ' + CAST(CAST(U_XX_BltWd as decimal(9,2)) AS VARCHAR)  ELSE CAST(U_XX_SNotes AS VARCHAR) END WHERE Code = '{0}'";

                using (var rs = new RecordsetAdapter(this.AddOn.Company))
                {
                    rs.Execute(string.Format(UpdateQuery, this.ControlManager.EditText("code").Value));
                }
                */
            }

        }


        private void UpdateControlBaseOnItemCode()
        {
            if (this.FormType == ConfigurationType.MetalDetector)
            {
                // < 110413 - Dom
                try
                {

                    var sp = new B1C.Utility.Helpers.StoredProcHelper(this.AddOn.Connection.Database + "_FTI.dbo.FTI_ConfigUpdateDetectorNotes");
                    sp.AddParameter("@code", System.Data.DbType.String, this.ControlManager.EditText("code").Value);

                    RecordsetAdapter.Execute(this.AddOn.Company, sp.ToString());
                }
                catch (Exception ex)
                {
                    this.AddOn.StatusBar.ErrorMessage = "[Dom] Failed to run stored procedure [FTI_ConfigUpdateDetectorNotes] to update system notes. Error:" + ex.Message;
                }

                // > 110413 - Dom

                /*  -- This is how Frank originally set it up, using a query.
                const string UpdateQuery = "UPDATE dbo.[@XX_METDET] SET U_XX_SNotes = CASE ISNULL(CAST(U_XX_SNotes as VARCHAR), '') WHEN '' THEN 'Model: ' + CAST(U_XX_Model as VARCHAR) ELSE CAST(U_XX_SNotes AS VARCHAR) END WHERE Code = '{0}'";

                using (var rs = new RecordsetAdapter(this.AddOn.Company))
                {
                    rs.Execute(string.Format(UpdateQuery, this.ControlManager.EditText("code").Value));
                }
                */
            }
            else
            {

                // < 110413 - Dom
                try
                {
                    var sp = new B1C.Utility.Helpers.StoredProcHelper(this.AddOn.Connection.Database + "_FTI.dbo.FTI_ConfigUpdateConveyorNotes");
                    sp.AddParameter("@code", System.Data.DbType.String, this.ControlManager.EditText("code").Value);

                    RecordsetAdapter.Execute(this.AddOn.Company, sp.ToString());

                }
                catch (Exception ex)
                {
                    this.AddOn.StatusBar.ErrorMessage = "[Dom] Failed to run stored procedure [FTI_ConfigUpdateConveyorNotes] tto update system notes. Error:" + ex.Message;
                }

                /*               
                const string UpdateQuery = "UPDATE dbo.[@XX_CONVBLT] SET U_XX_SNotes = CASE ISNULL(CAST(U_XX_SNotes as VARCHAR), '') WHEN '' THEN 'Size: ' + CAST(CAST(U_XX_OvLenght as decimal(9,2)) as varchar) +' X ' + CAST(CAST(U_XX_BltWd as decimal(9,2)) AS VARCHAR)  ELSE CAST(U_XX_SNotes AS VARCHAR) END WHERE Code = '{0}'";

                using (var rs = new RecordsetAdapter(this.AddOn.Company))
                {
                    rs.Execute(string.Format(UpdateQuery, this.ControlManager.EditText("code").Value));
                }
                */
            }

        }

        #endregion Methods
    }
}