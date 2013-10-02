//-----------------------------------------------------------------------
// <copyright file="SelectProductionOrders.cs" company="Fortress Technology Inc.">
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
    using SAPbouiCOM;

    #endregion Usings

    public class SelectProductionOrders : Window.FortressForm
    {
        #region Constuctors

        /// <summary>
        /// Initializes a new instance of the <see cref="SelectProductionOrders"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        public SelectProductionOrders(UI.AddOn addOn, string xmlPath, string formType) : base(addOn, xmlPath, formType) { }

        #endregion Constuctors

        #region Methods

        /// <summary>
        /// Initializes this instance.
        /// </summary>
        public override void Initialize()
        {
            this.ButtonClick += this.OnButtonClick;
            this.AfterArgumentsSet += this.OnAfterArgumentsSet;
            base.Initialize();
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public override void Dispose()
        {
            this.ButtonClick -= this.OnButtonClick;
            this.AfterArgumentsSet -= this.OnAfterArgumentsSet;
            base.Dispose();
        }


        /// <summary>
        /// Called when [after arguments set].
        /// </summary>
        /// <param name="sender">The sender.</param>
        private void OnAfterArgumentsSet(object sender)
        {
            this.Field("soNumber").Value = this.OrderId;
            this.Field("soLineNum").Value = this.OrderLine;

            this.SapForm.Title = string.Format("Production Orders for saler order {0} line {1}", this.OrderId, this.OrderLine);

            // Load Datatable and Matrix
            this.LoadProductionOrders();
        }



        /// <summary>
        /// Called when [button click].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.ItemEventArgs"/> instance containing the event data.</param>
        private void OnButtonClick(object sender, ItemEventArgs e)
        {
            int poNumber = this.GetSelectedPo();

            switch (e.ItemId)
            {
                case "viewOrd":
                    if (poNumber > 0)
                    {
                        this.OpenDocument(BoLinkedObject.lf_ProductionOrder, poNumber);
                    }
                    break;
            }
        }


        private int GetSelectedPo()
        {
            int selectedRow = -1;
            var matrix = this.ControlManager.Matrix("mtx_0");
            selectedRow = matrix.SelectedRow();

            if (selectedRow <= 0)
            {
                this.AddOn.MessageBox("No production order selected.");
                return -1;
            }

            return Int32.Parse(matrix[matrix.SelectedRow(), "V_-1"]);
        }


        public void LoadProductionOrders()
        {
            var statusTable = this.DataSource.UserTable("dtProdOr");
            statusTable.Rows.Clear();

            string query = string.Format(@"SELECT o.DocEntry, o.DueDate, o.ItemCode, i.ItemName FROM OWOR o JOIN OITM i on i.ItemCode = o.ItemCode WHERE o.OriginNum = {0} AND o.U_XX_SOLineNUm= {1}", this.OrderId, this.OrderLine);

            statusTable.ExecuteQuery(query);

            this.ControlManager.Matrix("mtx_0").BaseObject.LoadFromDataSource();

        }

        #endregion Methods
    }
}