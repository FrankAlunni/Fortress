using System;

namespace B1C.SAP.UI.Model.EventArguments
{
    public class ItemEventArgs : EventArgs
    {
        public string FormType { get; set; }
        public int FormMode { get; set; }
        public string FormId { get; set; }
        public string ItemId { get; set; }
        public string ColId { get; set; }
        public int Row { get; set; }
        public int KeyPressed { get; set; }
        public bool CancelEvent { get; set; }
        public string CancelMessage { get; set; }
        public bool ItemChanged { get; set; }

        public ItemEventArgs(string formType, int formMode, string formId, string itemId, string colId, int row)
        {
            this.FormType = formType;
            this.FormMode = formMode;
            this.FormId = formId;
            this.ItemId = itemId;
            this.ColId = colId;
            this.Row = row;
        }

        public ItemEventArgs(string formType, int formMode, string formId, string itemId, int keyPressed)
        {
            this.FormType = formType;
            this.FormMode = formMode;
            this.FormId = formId;
            this.ItemId = itemId;
            this.KeyPressed = keyPressed;
        }
    }
}
