using System;
using System.Collections.Generic;

using System.Text;
using SAPbobsCOM;
using B1C.SAP.DI.Helpers;

namespace B1C.SAP.DI.BusinessAdapters.Inventory
{
    public class PackageAdapter : SapObjectAdapter
    {
        private DocumentPackages _package;

        public DocumentPackages Current
        {
            get
            {
                return this._package;
            }
            private set
            {
                this._package = value;
            }
        }

        public PackageAdapter(Company company, DocumentPackages dp)
            : base(company)
        {
            this.Current = dp;
        }

        /// <summary>
        /// Sets the user defined field.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="fieldValue">The field value.</param>
        public void SetUserDefinedField(object fieldName, object fieldValue)
        {
            COMHelper.UserDefinedFieldValue(this._package.UserFields, fieldName, fieldValue);
        }

        /// <summary>
        /// Sets the user defined field.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>The Value of the field</returns>
        public object GetUserDefinedField(object fieldName)
        {
            return COMHelper.UserDefinedFieldValue(this._package.UserFields, fieldName);
        }

        protected override void Release()
        {
            if (this._package != null)
            {
                COMHelper.Release(ref this._package);
            }
        }
    }
}
