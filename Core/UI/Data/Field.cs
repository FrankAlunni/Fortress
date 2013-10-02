// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Field.cs" company="B1C Canada Inc.">
//   B1C Canada Inc.
// </copyright>
// <summary>
//   Defines the Field type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using B1C.SAP.DI.Business;

namespace B1C.SAP.UI.Data
{
    using Windows;

    /// <summary>
    /// Instance of a Field B1C.SAP
    /// </summary>
    public class Field
    {
        /// <summary>
        /// The current form
        /// </summary>
        private readonly Form _form;

        /// <summary>
        /// The field name
        /// </summary>
        private readonly string _fieldName;

        /// <summary>
        /// The source name
        /// </summary>
        private readonly string _sourceName;

        /// <summary>
        /// Initializes a new instance of the <see cref="Field"/> B1C.SAP.
        /// </summary>
        /// <param name="currentForm">The current form.</param>
        /// <param name="fieldname">The fieldname.</param>
        public Field(Form currentForm, string fieldname)
        {
            this._form = currentForm;
            this._fieldName = fieldname;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Field"/> B1C.SAP.
        /// </summary>
        /// <param name="currentForm">The current form.</param>
        /// <param name="sourcename">The sourcename.</param>
        /// <param name="fieldname">The fieldname.</param>
        public Field(Form currentForm, string sourcename, string fieldname)
        {
            this._form = currentForm;
            this._fieldName = fieldname;
            this._sourceName = sourcename;
        }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        /// <value>The value.</value>
        public string Value
        {
            get
            {
                if (string.IsNullOrEmpty(this._sourceName))
                {
                    return this._form.DataSource.UserSource(this._fieldName).Value.Trim(' ');
                }

                SAPbouiCOM.DBDataSource ds = this._form.DataSource.SystemTable(this._sourceName);
                if (ds == null)
                {
                    return string.Empty;
                }

                return ds.GetValue(this._fieldName, ds.Offset).Trim(' ');
            }

            set
            {
                if (string.IsNullOrEmpty(this._sourceName))
                {
                    this._form.DataSource.UserSource(this._fieldName).Value = value;
                    return;
                }

                SAPbouiCOM.DBDataSource ds = this._form.DataSource.SystemTable(this._sourceName);
                if (ds != null)
                {
                    ds.SetValue(this._fieldName, ds.Offset, value);
                }
            }
        }
    }
}
