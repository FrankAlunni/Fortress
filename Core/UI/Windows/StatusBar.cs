

namespace B1C.SAP.UI.Windows
{
    /// <summary>
    /// The Status Bar Class
    /// </summary>
    public class StatusBar
    {
        /// <summary>
        /// Private addon instance
        /// </summary>
        private AddOn addOn;

        /// <summary>
        /// Initializes a new instance of the <see cref="StatusBar"/> class.
        /// </summary>
        /// <param name="addon">The addon.</param>
        public StatusBar(AddOn addon)
        {
            this.addOn = addon;
        }

        /// <summary>
        /// Sets the error message.
        /// </summary>
        /// <value>The error message.</value>
        public string ErrorMessage
        {
            set
            {
                this.addOn.Application.SetStatusBarMessage(value, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// Sets the message.
        /// </summary>
        /// <value>The message.</value>
        public string Message
        {
            set
            {
                this.addOn.Application.SetStatusBarMessage(value, SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
        }
    }
}
