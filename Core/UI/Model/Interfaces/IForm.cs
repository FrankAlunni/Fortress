// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IForm.cs" company="B1C Canada Inc.">
//   B1C Canada Inc.
// </copyright>
// <summary>
//   Defines the IForm type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using SAPbouiCOM;

namespace B1C.SAP.UI.Model.Interfaces
{
    using Windows;

    /// <summary>
    /// IForm interface
    /// </summary>
    public interface IForm
    {
        /// <summary>
        /// Opens the document.
        /// </summary>
        /// <param name="documentType">Type of the document.</param>
        /// <param name="documentEntry">The document entry.</param>
        void OpenDocument(SAPbouiCOM.BoLinkedObject documentType, int documentEntry);

        /// <summary>
        /// Freezes the specified is frozen.
        /// </summary>
        /// <param name="isFrozen">if set to <c>true</c> [is frozen].</param>
        void Freeze(bool isFrozen);

        /// <summary>
        /// Refreshes this instance.
        /// </summary>
        void Refresh();

        /// <summary>
        /// Applies the filters.
        /// </summary>
        void ApplyFilters();

        /// <summary>
        /// Gets the add on.
        /// </summary>
        /// <value>The add on.</value>
        AddOn AddOn { get; }

        /// <summary>
        /// Occurs when [disposed].
        /// </summary>
        event DisposeHandler Disposed;

        /// <summary>
        /// Gets the sap form.
        /// </summary>
        /// <value>The sap form.</value>
        SAPbouiCOM.Form SapForm { get; }

        /// <summary>
        /// Gets the form id.
        /// </summary>
        /// <value>The form id.</value>
        string FormId { get; }

        /// <summary>
        /// Gets the form type id.
        /// </summary>
        /// <value>The form id.</value>
        string FormTypeId { get; }

        /// <summary>
        /// Tools the bar clicked.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        /// <param name="beforeEvent">if set to <c>true</c> [before event].</param>
        /// <param name="cancelEvent">if set to <c>true</c> [cancel event].</param>
        /// <param name="cancelMessage">The cancel message.</param>
        void ToolBarClicked(string buttonId, bool beforeEvent, out bool cancelEvent, out string cancelMessage);

        /// <summary>
        /// Menu was clicked.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        /// <param name="beforeEvent">if set to <c>true</c> [before event].</param>
        /// <param name="cancelEvent">if set to <c>true</c> [cancel event].</param>
        /// <param name="cancelMessage">The cancel message.</param>
        void MenuClicked(string buttonId, bool beforeEvent, out bool cancelEvent, out string cancelMessage);

        /// <summary>
        /// Rights the clicked.
        /// </summary>
        /// <param name="contextMenuInfo">The context menu info.</param>
        /// <param name="bubbleEvent">if set to <c>true</c> [bubble event].</param>
        void RightClicked(ref ContextMenuInfo contextMenuInfo, out bool bubbleEvent);

        /// <summary>
        /// Initializes this instance.
        /// </summary>
        void Initialize();

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources
        /// </summary>
        void Dispose();
    }
}
