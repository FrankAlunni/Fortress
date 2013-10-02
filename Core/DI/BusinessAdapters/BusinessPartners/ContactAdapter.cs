// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ContactAdapter.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the ContactAdapter type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using B1C.SAP.DI.Helpers;

namespace B1C.SAP.DI.BusinessAdapters.BusinessPartners
{
    using SAPbobsCOM;
    using Utility.Helpers;

    /// <summary>
    /// The Contact (Activity) adapter
    /// </summary>
    public class ContactAdapter : SapObjectAdapter
    {
        /// <summary>
        /// The contact
        /// </summary>
        private Contacts contact;

        /// <summary>
        /// Initializes a new instance of the <see cref="ContactAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="contact">The contact.</param>
        public ContactAdapter(Company company, Contacts contact)
            : base(company)
        {
            this.Contact = contact;
        }

        /// <summary>
        /// Gets the contact.
        /// </summary>
        /// <value>The contact.</value>
        public Contacts Contact
        {
            get { return this.contact; }
            private set { this.contact = value; }
        }

        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected override void Release()
        {
            if (this.Contact != null)
            {
                COMHelper.Release(ref this.contact);
            }
        }

        /// <summary>
        /// Sets the user defined field.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="fieldValue">The field value.</param>
        public void SetUserDefinedField(object fieldName, object fieldValue)
        {
            COMHelper.UserDefinedFieldValue(this.contact.UserFields, fieldName, fieldValue);
        }

        /// <summary>
        /// Gets the user defined field.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>The Field Value</returns>
        public object GetUserDefinedField(object fieldName)
        {
            return COMHelper.UserDefinedFieldValue(this.contact.UserFields, fieldName);
        }
    }
}
