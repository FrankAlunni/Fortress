// --------------------------------------------------------------------------------------------------------------------
// <copyright file="PaymentAdapter.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the PaymentAdapter type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using B1C.SAP.DI.Helpers;

namespace B1C.SAP.DI.BusinessAdapters.Banking
{
    using SAPbobsCOM;

    /// <summary>
    /// The Payment Adapter
    /// </summary>
    public class PaymentAdapter : SapObjectAdapter
    {
        /// <summary>
        /// The payment.
        /// </summary>
        private Payments payment;

        /// <summary>
        /// Initializes a new instance of the <see cref="PaymentAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="payment">The payment.</param>
        public PaymentAdapter(Company company, Payments payment)
            : base(company)
        {
            this.Payment = payment;
        }

        /// <summary>
        /// Gets the payment.
        /// </summary>
        /// <value>The payment.</value>
        public Payments Payment
        {
            get
            {
                return this.payment;
            }

            private set
            {
                this.payment = value;
            }
        }

        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected override void Release()
        {
            COMHelper.Release(ref this.payment);
        }

        /// <summary>
        /// Sets the user defined field.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="fieldValue">The field value.</param>
        public void SetUserDefinedField(object fieldName, object fieldValue)
        {
            COMHelper.UserDefinedFieldValue(this.payment.UserFields, fieldName, fieldValue);
        }

        /// <summary>
        /// Sets the user defined field.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>The Value of the field</returns>
        public object GetUserDefinedField(object fieldName)
        {
            return COMHelper.UserDefinedFieldValue(this.payment.UserFields, fieldName);
        }
    }
}
