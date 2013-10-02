//-----------------------------------------------------------------------
// <copyright file="ConveyorConfigurator.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
// <email>frank.alunni@b1computing.com</email>
//-----------------------------------------------------------------------

namespace B1C.SAP.Addons.Configurator.Forms
{
    #region Usings
    using UI.Model.EventArguments;
    using System;
    using DI.BusinessAdapters;
    using DI.Business;
    #endregion Usings

    public class ProductConfigurator : Window.FortressConfiguration 
    {
        #region Constuctors
        /// <summary>
        /// Initializes a new instance of the <see cref="ProductConfigurator"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        public ProductConfigurator(UI.AddOn addOn, string xmlPath, string formType) : base(addOn, xmlPath, formType) { }
        #endregion

        #region Methods
        /// <summary>
        /// Initializes this instance.
        /// </summary>
        public override void Initialize()
        {
            FolderClick += this.OnFolderClick;
            AfterFormLoad += this.OnAfterFormLoad;
            AfterArgumentsSet += this.OnAfterArgumentsSet;
            base.Initialize();
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public override void Dispose()
        {
            FolderClick -= this.OnFolderClick;
            AfterFormLoad -= this.OnAfterFormLoad;
            AfterArgumentsSet -= this.OnAfterArgumentsSet;
            base.Dispose();
        }

        public override Configurator.Window.FortressConfiguration.ConfigurationType FormType
        {
            get { return ConfigurationType.MetalDetector; }
        }


        /// <summary>
        /// Called when [after arguments set].
        /// </summary>
        /// <param name="sender">The sender.</param>
        private void OnAfterArgumentsSet(object sender)
        {
            if (this.IsNewRecord)
            {
                this.FindCode = this.OrderId + '-' + this.OrderLine;

                var loadData = new System.Text.StringBuilder();
                loadData.AppendLine("<DataLoader UDOName='MDConfigurator' UDOType='MasterData'>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", this.FindCode);
                loadData.AppendFormat("<Field Name='U_XX_OrderNo' Type='System.Int32' Value='{0}' />", this.OrderId);
                loadData.AppendFormat("<Field Name='U_XX_OrdrLnNo' Type='System.Int32' Value='{0}' />", this.OrderLine);
                loadData.AppendFormat("<Field Name='U_XX_Revis' Type='System.String' Value='{0}' />", 1);
                loadData.AppendFormat("<Field Name='U_XX_Status' Type='System.String' Value='{0}' />", "Initiated");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("</DataLoader>");

                var loader = DataLoader.InstanceFromXMLString(loadData.ToString());
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
                this.FindCode = string.Empty;

                // Determine the Code to find
                using (var rs = new RecordsetAdapter(this.AddOn.Company, string.Format("select Code from [@XX_METDET] WHERE U_XX_OrderNo = '{0}' AND U_XX_OrdrLnNo = {1}", this.OrderId, this.OrderLine)))
                {
                    if (!rs.EoF)
                    {
                        this.FindCode = rs.FieldValue("Code").ToString();
                    }
                }

            }
        }

        /// <summary>
        /// Raises the AfterFormLoad event.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        private void OnAfterFormLoad(object sender, FormEventArgs e)
        {
            try
            {
                var folder = (SAPbouiCOM.Folder)this.ControlManager.Item("tabGen").BaseItem.Specific;
                folder.Select();
                folder = (SAPbouiCOM.Folder)this.ControlManager.Item("tabProdInM").BaseItem.Specific;
                folder.Select();
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the FolderClick event.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        private void OnFolderClick(object sender, ItemEventArgs e)
        {
            var folder = (SAPbouiCOM.Folder)this.ControlManager.Item(e.ItemId).BaseItem.Specific;
            this.SapForm.PaneLevel = int.Parse(folder.ValOn);
            if (!folder.Selected)
            {
                folder.Select();
            }

            switch (e.ItemId)
            {
                case "tabGen":
                    folder = (SAPbouiCOM.Folder)this.ControlManager.Item("tabProdInM").BaseItem.Specific;
                    folder.Select();
                    break;
                case "tabMetalD":
                    folder = (SAPbouiCOM.Folder)this.ControlManager.Item("tabMetalDM").BaseItem.Specific;
                    folder.Select();
                    break;
            }
        }
        #endregion Methods

    }
}
