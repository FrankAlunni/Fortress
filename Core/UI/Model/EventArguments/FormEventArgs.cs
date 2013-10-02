// --------------------------------------------------------------------------------------------------------------------
// <copyright file="FormEventArgs.cs" company="B1C Canada Inc.">
//   B1C Canada Inc.
// </copyright>
// <summary>
//   Defines the FormEventArgs type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.UI.Model.EventArguments
{
    using System;

    /// <summary>
    /// Forms event arguments
    /// </summary>
    public class FormEventArgs : EventArgs
    {
        public string FormType { get; set; }
        public string FormId { get; set; }
        public bool CancelEvent { get; set; }
        public string CancelMessage { get; set; }
        public string ObjectXml { get; set; }
        public bool IsSuccessful { get; internal set; }

        public FormEventArgs(string formType, string formId, string objectXml)
        {
            this.FormType = formType;
            this.FormId = formId;
            this.ObjectXml = objectXml;

            this.IsSuccessful = true;
        }
    }
}
