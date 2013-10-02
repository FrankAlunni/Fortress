// <copyright file="FormulaException.cs" company="B1C Canada Inc.">
//   B1C Canada Inc.
// </copyright>
// <author>Bryan Atkinson</author>
// <summary>
//   Defines the DeliveryNoteListener type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.Model.Exceptions
{

    #region Using Directive(s)
    
    using System;    

    #endregion Using Directive(s)

    public class FormulaException : ApplicationException
    {

        private Exception _innerException;

        /// <summary>
        /// The Partner Code
        /// </summary>
        public string PartnerCode
        {
            get;
            private set;
        }

        /// <summary>
        /// The Item Code
        /// </summary>
        public string ItemCode
        {
            get;
            private set;
        }

        /// <summary>
        /// The Formula Code
        /// </summary>
        public string FormulaCode
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets a message that describes the current exception.
        /// </summary>
        /// <value></value>
        /// <returns>The error message that explains the reason for the exception, or an empty string("").</returns>
        public override string Message
        {
            get
            {
                string baseMessage = string.Empty;
                if (this._innerException != null)
                {
                    baseMessage = this._innerException.Message.Substring(this._innerException.Message.LastIndexOf(':') + 1).Trim();
                }
                else
                {
                    baseMessage = "Unknown Error";
                }

                return string.Format("Error in Formula {0}, for Partner {1} on Item {2}: {3}", this.FormulaCode, this.PartnerCode, this.ItemCode, baseMessage);
            }
        }

        /// <summary>
        /// Initializes a new instance of  <see cref="FormulaException"/> class.
        /// </summary>
        /// <param name="partnerCode">The partner code</param>
        /// <param name="itemCode">The item code</param>
        /// <param name="formulaCode">The formula code</param>
        /// <param name="ex">The inner exception</param>
        public FormulaException(Exception ex, string partnerCode, string itemCode, string formulaCode)
        {
            this._innerException = ex;
            this.PartnerCode = partnerCode;
            this.ItemCode = itemCode;
            this.FormulaCode = formulaCode;
        }
    }
}
