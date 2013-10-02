using System;

namespace B1C.SAP.UI.Model.EventArguments
{
    public class MenuEventArgs : EventArgs
    {
        public string MenuId { get; set; }

        public MenuEventArgs(string menuId)
        {
            this.MenuId = menuId;
        }
    }
}
