using System;
using SAPbouiCOM;

namespace B1C.SAP.UI.Model.EventArguments
{
    public class ChooseFromListEventArgs : EventArgs
    {
        public string FormType { get; set; }
        public int FormMode { get; set; }
        public string FormId { get; set; }
        public string ItemId { get; set; }
        public DataTable SelectedObjects { get; set; }
        public bool CancelEvent { get; set; }
        public int SelectedRow { get; set; }
        public string SelectedColUID { get; set; }

        public ChooseFromListEventArgs(string formType, int formMode, string formId, string itemId, DataTable selectedObjects, int row, string colUid)
        {
            this.FormType = formType;
            this.FormMode = formMode;
            this.FormId = formId;
            this.ItemId = itemId;
            this.SelectedObjects = selectedObjects;
            this.SelectedRow = row;
            this.SelectedColUID = colUid;
        }
    }
}
