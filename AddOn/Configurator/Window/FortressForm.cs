//-----------------------------------------------------------------------
// <copyright file="FortressForm.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
// <email>frank.alunni@b1computing.com</email>
//-----------------------------------------------------------------------

namespace B1C.SAP.Addons.Configurator.Window
{
    public class FortressForm : UI.Windows.Form
    {
        #region Constuctors

        /// <summary>
        /// Initializes a new instance of the <see cref="FortressForm"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        public FortressForm(UI.AddOn addOn, string xmlPath, string formType) : base(addOn, xmlPath, formType) { }

        #endregion Constuctors

        #region Properties

        /// <summary>
        /// Gets the document number.
        /// </summary>
        /// <value>The document number.</value>
        public new string DocumentNumber
        {
            get
            {
                if (this.Arguments == null)
                {
                    return "0";
                }

                object argument;
                this.Arguments.TryGetValue("DocumentNumber", out argument);
                if (argument != null)
                {
                    return argument.ToString();
                }

                return "0";
            }
        }

        /// <summary>
        /// Gets the order id.
        /// </summary>
        /// <value>The order id.</value>
        protected internal string OrderId
        {
            get
            {
                if (this.Arguments == null)
                {
                    return "0";
                }

                object argument;
                this.Arguments.TryGetValue("OrderId", out argument);
                if (argument != null)
                {
                    return argument.ToString();
                }

                return "0";
            }
        }

        /// <summary>
        /// Gets the order line.
        /// </summary>
        /// <value>The order line.</value>
        protected internal string OrderLine
        {
            get
            {
                if (this.Arguments == null)
                {
                    return "0";
                }

                object argument;
                this.Arguments.TryGetValue("OrderLine", out argument);
                if (argument != null)
                {
                    return argument.ToString();
                }

                return "0";
            }
        }

        /// <summary>
        /// Gets the order line.
        /// </summary>
        /// <value>The order line.</value>
        protected string RowNumber
        {
            get
            {
                if (this.Arguments == null)
                {
                    return "0";
                }

                object argument;
                this.Arguments.TryGetValue("RowNumber", out argument);
                if (argument != null)
                {
                    return argument.ToString();
                }

                return "0";
            }
        }

        /// <summary>
        /// Gets the order line.
        /// </summary>
        /// <value>The order line.</value>
        protected string ItemCode
        {
            get
            {
                if (this.Arguments == null)
                {
                    return "";
                }

                object argument;
                this.Arguments.TryGetValue("ItemCode", out argument);
                if (argument != null)
                {
                    return argument.ToString();
                }

                return "";
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is new record.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is new record; otherwise, <c>false</c>.
        /// </value>
        protected bool IsNewRecord
        {
            get
            {
                object argument;
                this.Arguments.TryGetValue("RecordExists", out argument);
                if (argument != null)
                {
                    return !bool.Parse(argument.ToString());
                }

                return false;
            }
        }

        #endregion Properties
    }
}