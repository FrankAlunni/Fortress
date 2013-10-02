using System;
using System.Collections.Generic;

using System.Text;
using SAPbobsCOM;

namespace B1C.SAP.DI.Model.Exceptions
{
    [Serializable()]
    public class SapException : ApplicationException
    {

        /// <summary>
        /// Gets or sets the error code.
        /// </summary>
        /// <value>The error code.</value>
        public int ErrorCode
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets or sets the error message.
        /// </summary>
        /// <value>The error message.</value>
        public string ErrorMessage
        {
            get;
            private set;
        }


        /// <summary>
        /// Gets a value indicating whether or not to retry this operation.
        /// This normally indicates that SAP is busy, or cannot perform the operation at this time.
        /// </summary>
        /// <value><c>true</c> if a retry is necessary; otherwise, <c>false</c>.</value>
        public bool DoRetry
        {
            get 
            { 
                bool doRetry = false;

                // TODO: Make this configuration/table driven

                // ErrorCode -2038 indicates that there is blocking going on on the SAP DB
                if (Math.Abs(this.ErrorCode) == 2038)
                {
                    doRetry = true;
                }

                return doRetry;
            }
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
                return string.Format("SAP Business rule error. No:[{0}] - {1}. {2}", this.ErrorCode, this.ErrorMessage, base.Message);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SapException"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        public SapException(Company company)
        {
            int errorCode;
            string errorMessage;

            company.GetLastError(out errorCode, out errorMessage);

            this.ErrorCode = errorCode;
            this.ErrorMessage = errorMessage;
        }
    }
}
