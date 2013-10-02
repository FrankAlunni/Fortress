//-----------------------------------------------------------------------
// <copyright file="ProductionOrder.cs" company="Fortress Technology Inc.">
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

    public class ProductionOrder : UI.Windows.Form
    {
        private int selectedRow = -1;

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="UI.Windows.Form"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        public ProductionOrder(UI.AddOn addOn)
            : base(addOn)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UI.Windows.Form"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="form">The SAP form.</param>
        public ProductionOrder(UI.AddOn addOn, Form form)
            : base(addOn, form)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UI.Windows.Form"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        public ProductionOrder(UI.AddOn addOn, string xmlPath, string formType)
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
        public ProductionOrder(UI.AddOn addOn, string xmlPath, string formType, string objectType)
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
            ButtonClick += this.OnButtonClick;
            FormResize += this.OnFormResize;
            AfterItemClick += this.OnAfterItemClick;
        }


        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public override void Dispose()
        {
            base.Dispose();
            ButtonClick -= this.OnButtonClick;
            FormResize -= this.OnFormResize;
            AfterItemClick -= this.OnAfterItemClick;
        }

        private void OnAfterItemClick(object sender, ItemEventArgs e)
        {
            if (e.ItemId == "37")
            {
                selectedRow = e.Row;
                this.DisableButtons();
            }
        }

        /// <summary>
        /// Creates the production order for the selected line.
        /// </summary>
        private void DisableButtons()
        {
            bool enableCreate = false;

            int docEntry = int.Parse(this.PrimaryKey);
            if (docEntry > 0)
            {
                using (var adapter = new FortressProductionOrderAdapter(this.AddOn.Company, docEntry))
                {
                    int currentLine = this.GetSelectedLineNum();
                    if (currentLine >= 0)
                    {
                        enableCreate = !adapter.HasProductionOrder(currentLine);
                    }
                }
            }

            this.ControlManager.Item("prodord").Enabled = enableCreate;

        }

        private int GetSelectedLineNum()
        {
            if (selectedRow < 0)
            {
                throw new FortressException("Production order detail selected.");
            }

            if (this.DataSource.SystemTable("WOR1").Size <= (selectedRow - 1)) return -1;
            return int.Parse(this.DataSource.SystemTable("WOR1").GetValue("LineNum", selectedRow - 1));
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

            UI.Adapters.Item createProdOrderBtn = this.ControlManager.Item("prodord");
            createProdOrderBtn.MoveOnRightOf(cancelBtn, 5);

        }

        /// <summary>
        /// Raises the AfterItemClick event.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        protected void OnButtonClick(object sender, ItemEventArgs e)
        {
            if (e.ItemId != "prodord") return;

            this.SapForm.Items.Item(e.ItemId).Enabled = false;

            try
            {
                // Create PO
                this.CreateProductionOrder();
            }
            finally
            {
                this.SapForm.Items.Item(e.ItemId).Enabled = true;
            }
        }

        protected override string PrimaryKey
        {
            get
            {
                try
                {
                    string fullKey = this.SapForm.BusinessObject.Key;

                    fullKey = fullKey.Substring(fullKey.IndexOf("AbsoluteEntry>") + 14);
                    fullKey = fullKey.Substring(0, fullKey.IndexOf("<"));

                    return fullKey;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// Creates the production order for the selected line.
        /// </summary>
        public void CreateProductionOrder()
        {
            int docEntry = int.Parse(this.PrimaryKey);

            if (docEntry <= 0)
            {
                throw new FortressException(string.Format("Can't load current production order {0} for production order creation. Has the current production order been saved?", docEntry));
            }

            // Load the sales order
            using (var adapter = new FortressProductionOrderAdapter(this.AddOn.Company, docEntry))
            {
                int currentLine = this.GetSelectedLineNum();
                if (currentLine < 0)
                {
                    this.AddOn.StatusBar.ErrorMessage = "No production order detail selected.";
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
        #endregion Method(s)
    }
}