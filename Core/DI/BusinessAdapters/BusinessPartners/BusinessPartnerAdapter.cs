//-----------------------------------------------------------------------
// <copyright file="BusinessPartnerAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.DI.BusinessAdapters.BusinessPartners
{
    using System;
    using System.Collections.Generic;
    using Components;
    using Helpers;
    using SAPbobsCOM;
    using B1C.SAP.DI.Model.Exceptions;

    /// <summary>
    /// Instance of the BusinessPartnerAdapter class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public class BusinessPartnerAdapter : SapObjectAdapter
    {
        #region Fields
        /// <summary>
        /// The SAP BusinessPartners document object
        /// </summary>
        private BusinessPartners document;

        /// <summary>
        /// Collection of addresses
        /// </summary>
        private AddressLines lines;
        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="BusinessPartnerAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP company object</param>
        public BusinessPartnerAdapter(Company company) : base(company)
        {
            this.Document = (BusinessPartners)this.Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="BusinessPartnerAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object.</param>
        /// <param name="cardCode">The card code.</param>
        public BusinessPartnerAdapter(Company company, string cardCode)
            : this(company)
        {
            this.Load(cardCode);
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets or sets the document.
        /// </summary>
        /// <value>The document.</value>
        public BusinessPartners Document
        {
            get { return this.document; }

            protected set { this.document = value; }
        }

        /// <summary>
        /// Gets the lines.
        /// </summary>
        /// <value>The lines.</value>
        public AddressLines Lines
        {
            get
            {
                if (this.document == null)
                {
                    return null;
                }

                if (this.lines == null)
                {
                    this.lines = new AddressLines(this.document.Addresses);
                }

                return this.lines;
            }
        }

        /// <summary>
        /// Gets the shipping address.
        /// </summary>
        /// <value>The shipping address.</value>
        public BPAddresses ShippingAddress
        {
            get
            {
                foreach (BPAddresses line in this.Lines)
                {
                    if (line.AddressType == BoAddressType.bo_ShipTo)
                    {
                        return this.lines.Current;
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the billing address.
        /// </summary>
        /// <value>The billing address.</value>
        public BPAddresses BillingAddress
        {
            get
            {
                if (this.Lines == null)
                {
                    return null;    
                }

                foreach (BPAddresses line in this.Lines)
                {
                    if (line.AddressType == BoAddressType.bo_BillTo)
                    {
                        return this.Lines.Current;
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the currency.
        /// </summary>
        /// <value>The currency.</value>
        public string Currency
        {
            get
            {
                return this.Document.Currency;
            }
        }
        #endregion Properties

        #region Methods
        /// <summary>
        /// Instantiates a Business Partner Adapter with the given company
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <returns>An instance of the BusinessPartnerAdapter class</returns>
        public static BusinessPartnerAdapter Instance(Company company)
        {
            return new BusinessPartnerAdapter(company);
        }

        /// <summary>
        /// Instances the specified parent add on.
        /// </summary>
        /// <param name="company">The SAP company object.</param>
        /// <param name="cardCode">The card code.</param>
        /// <returns>An instance to the BusinessPartnerAdapter class</returns>
        public static BusinessPartnerAdapter Instance(Company company, string cardCode)
        {
            return new BusinessPartnerAdapter(company, cardCode);
        }

        /// <summary>
        /// Loads the specified card code.
        /// </summary>
        /// <param name="cardCode">The card code.</param>
        /// <returns>True if the method loads successfully</returns>
        public bool Load(string cardCode)
        {
            if (this.Company == null)
            {
                throw new ArgumentNullException("cardCode");
            }

            return this.Document.GetByKey(cardCode);
        }

        /// <summary>
        /// Returns true if the business partner is a program
        /// </summary>
        /// <returns>
        /// True if the business partner is a program
        /// </returns>
        public bool IsProgram()
        {
            int groupCode;
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "ProgramGroupId", out groupCode);

            using (var rows = new RecordsetAdapter(this.Company, string.Format("SELECT * FROM [OCRD] WHERE CardCode = '{0}' AND GroupCode = {1}", this.Document.CardCode, groupCode)))
            {
                return !rows.EoF;
            }
        }

        /// <summary>
        /// Returns true if the business partner is a vendor
        /// </summary>
        /// <returns>
        /// True if the business partner is a vendor
        /// </returns>
        public bool IsVendor()
        {
            using (var rows = new RecordsetAdapter(this.Company, string.Format("SELECT * FROM [OCRD] WHERE CardCode = '{0}' AND CardType = 'S'", this.Document.CardCode)))
            {
                return !rows.EoF;
            }
        }

        /// <summary>
        /// Adds this instance.
        /// </summary>
        /// <returns>True if the method adds successfully</returns>
        public bool Add()
        {
            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the BusinessPartnerAdapter.Add is not set.");
            }

            if (this.Document == null)
            {
                throw new NullReferenceException("The document has not been set. BusinessPartnerAdapter.Add");
            }

            if (this.Document.Add() != 0)
            {
                throw new SapException(this.Company);
            }

            return true;
        }

        /// <summary>
        /// Updates this instance.
        /// </summary>
        /// <returns>True if the method updates successfully</returns>
        public bool Update()
        {
            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the BusinessPartnerAdapter.Update is not set.");
            }

            if (this.Document == null)
            {
                throw new NullReferenceException("The document has not been set. BusinessPartnerAdapter.Update");
            }

            if (this.Document.Update() != 0)
            {
                throw new SapException(this.Company);
            }

            return true;
        }

        /// <summary>
        /// Check if the item is valid for usage from a business partner
        /// </summary>
        /// <param name="itemCode">The item code.</param>
        /// <returns>Blank if valid, otherwise the reason why is not valid</returns>
        public string IsItemValidForUsage(string itemCode)
        {
            var item = new Inventory.ItemAdapter(this.Company, itemCode);

            if (this.IsProgram())
            {
                if (item.IsExclusive())
                {
                    if (!item.IsPartnerExclusive(this.Document.CardCode))
                    {
                        return string.Format("The item {0} is exclusive to another program.", itemCode);
                    }
                }
            }

            return string.Empty;
        }

        /// <summary>Gets a list of the subvendors</summary>
        /// <returns>The list of the subvendors</returns>
        public Dictionary<string, string> SubVendors()
        {
            var subvendors = new Dictionary<string, string>();
            string query = string.Format("SELECT T1.CardCode, T1.CardName FROM [@XX_DISTRIBUTOR] T0 JOIN OCRD T1 ON T0.U_Vendor = T1.CardCode WHERE T0.U_Distrib = '{0}'", this.Document.CardCode);

            using (var rows  = new RecordsetAdapter(this.Company, query))
            {
                while (!rows.EoF)
                {
                    subvendors.Add(rows.FieldValue("CardCode").ToString(), rows.FieldValue("CardName").ToString());
                    rows.MoveNext();
                }
            }

            return subvendors;
        }

        /// <summary>Gets a list of the subvendors</summary>
        /// <returns>The list of the subvendors</returns>
        public bool HasSubVendors()
        {
            string query = string.Format("SELECT T1.CardCode, T1.CardName FROM [@XX_DISTRIBUTOR] T0 JOIN OCRD T1 ON T0.U_Vendor = T1.CardCode WHERE T0.U_Distrib = '{0}'", this.Document.CardCode);

            using (var rows = new RecordsetAdapter(this.Company, query))
            {
                return !rows.EoF;
            }
        }

        /// <summary>
        /// Check if the current vendor is a one time vendor
        /// </summary>
        /// <returns>true, otherwise false</returns>
        public bool IsOneTimeVendor()
        {
            string oneTimeVendorCode;
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "OneTimeVendorCode", out oneTimeVendorCode);
            if (string.IsNullOrEmpty(oneTimeVendorCode))
            {
                return false;
            }

            return this.Document.CardCode == oneTimeVendorCode;
        }

        /// <summary>
        /// Sets the user defined field on the current line.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <param name="fieldValue">The field value.</param>
        public void SetUserDefinedField(object fieldIndex, object fieldValue)
        {
            COMHelper.UserDefinedFieldValue(this.Document.UserFields, fieldIndex, fieldValue);
        }

        /// <summary>
        /// Gets the user defined field from the Document.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <returns>The value of the field</returns>
        public object GetUserDefinedField(object fieldIndex)
        {
            return COMHelper.UserDefinedFieldValue(this.Document.UserFields, fieldIndex);
        }

        /// <summary>
        /// Gets the default payment method for this business partner.
        /// </summary>
        /// <returns>The default payment method's PaymentMethodCode, or an empty string if there is no default</returns>
        public string DefaultPaymentMethod
        {
            get
            {
                string paymentMethod = string.Empty;

                BPPaymentMethods bppm = this.Document.BPPaymentMethods;
                try
                {
                    if (bppm != null && bppm.Count > 0)
                    {
                        bppm.SetCurrentLine(0);
                        paymentMethod = bppm.PaymentMethodCode;
                    }
                }
                finally
                {
                    COMHelper.Release(ref bppm);
                }

                return paymentMethod;
            }
        }

        /// <summary>
        /// Releases the Documents COM Object.
        /// </summary>
        protected override void Release()
        {
            if (this.lines != null)
            {
                this.lines.Dispose();
            }

            if (this.document != null)
            {
                COMHelper.Release(ref this.document);
            }
        }
        #endregion Methods
    }
}
