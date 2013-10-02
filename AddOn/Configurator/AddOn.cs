//-----------------------------------------------------------------------
// <copyright file="AddOn.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
// <email>frank.alunni@b1computing.com</email>
//-----------------------------------------------------------------------

namespace B1C.SAP.Addons.Configurator
{
    #region Usings

    using DI.Constants;
    using UI.Model.EventArguments;
    using Forms;
    using SAPbobsCOM;
    using DI.Model.EventArguments;
    using Model;
    using Fortress.DI.Model.Exceptions;

    #endregion

    /// <summary>
    /// The instance of the addon class
    /// </summary>
    public class AddOn : UI.AddOn
    {
        #region Constructors
        /// <summary>
        ///   Initializes a new instance of the <see cref = "AddOn" /> class.
        /// </summary>
        /// <param name = "addonNamespace">The addon namespace.</param>
        /// <param name = "addonId">The addon id.</param>
        public AddOn(string addonNamespace, string addonId) : base(addonNamespace, addonId)
        {
            this.CaptureEvents(SapDocumentTables.GetFormCode(BoObjectTypes.oOrders).ToString()); 
            //CaptureEvents(systemFormType);
            // RegisterEvents();
        }
        #endregion Constructors

        #region Methods
        /// <summary>
        ///   Releases the events.
        /// </summary>
        public override void UnregisterEvents()
        {
            NewFormCreated -= OnNewFormCreated;
            UpdateDatabase -= OnUpdateDatabase;
        }

        /// <summary>
        ///   Handles the events.
        /// </summary>
        public override void RegisterEvents()
        {
            NewFormCreated += OnNewFormCreated;
            UpdateDatabase += OnUpdateDatabase;
        }

        /// <summary>
        /// Handles the exception.
        /// </summary>
        /// <param name="ex">The ex.</param>
        public override void HandleException(System.Exception ex)
        {
            if (ex is FortressException)
            {
                this.StatusBar.ErrorMessage = ex.Message;
            }
            else
            {
                base.HandleException(ex);
            }
        }

        /// <summary>
        /// Called when [update database].
        /// </summary>
        /// <param name="sender">The sender.</param>
        private void OnUpdateDatabase(object sender)
        {
            if (!DatabaseUpdater.CreateCommission(this))
            {
                return;
            }

            if (!DatabaseUpdater.CreateApproval(this))
            {
                return;
            }

            if (!DatabaseUpdater.CreateConveyorBeltConfigurator(this))
            {
                return;
            }

            if (!DatabaseUpdater.CreateMetalDetectorConfigurator(this))
            {
                return;
            }

            if (!DatabaseUpdater.CreateGeneralSettings(this))
            {
                return;
            }

            if (!DatabaseUpdater.UpdateSalesOrder(this))
            {
                return;
            }
        }

        /// <summary>
        ///   Called when [new form created].
        /// </summary>
        /// <param name = "sender">The sender.</param>
        /// <param name = "e">The <see cref = "B1C.SAP.UI.Model.EventArguments.FormEventArgs" /> instance containing the event data.</param>
        private void OnNewFormCreated(object sender, FormEventArgs e)
        {
            // Add any new system form to the form collection
            // This is used to capture when a new system form is created
            // the event will be raised only for forms there the CaptureEvents
            // has been called in the constructor fo this class
            if (e.FormType == SapDocumentTables.GetFormCode(BoObjectTypes.oOrders).ToString()) 
            { 
                FormFactory.Create<SalesOrder>(Application.Forms.Item(e.FormId)); 
            }
            else if (e.FormType == SapDocumentTables.GetFormCode(BoObjectTypes.oProductionOrders).ToString())
            {
                FormFactory.Create<ProductionOrder>(Application.Forms.Item(e.FormId));
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="StatusbarChangedEventArgs"/> instance containing the event data.</param>
        internal void StatusbarChanged(object sender, StatusbarChangedEventArgs e)
        {
            if (e.IsErrorMessage)
            {
                this.StatusBar.ErrorMessage = e.ProgressMessage;
            }
            else
            {
                this.StatusBar.Message = e.ProgressMessage;
            }
        }

        #endregion
    }
}
