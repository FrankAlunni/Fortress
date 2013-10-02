//-----------------------------------------------------------------------
// <copyright file="ProductionOrderAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.Production
{
    #region Using Directive(s)

    using SAPbobsCOM;
    using Helpers;
    using System;
    using Model.Exceptions;
    using Utility.Logging;

    #endregion Using Directive(s)

    public class ProductionOrderAdapter : SapObjectAdapter
    {
        #region Private Member(s)

        private ProductionOrderLines _lines;

        #endregion Private Member(s)

        #region Properties

        /// <summary>
        /// Gets a value indicating whether this <see cref="ProductionOrderAdapter"/> exists.
        /// </summary>
        /// <value>
        ///   <c>true</c> if exists; otherwise, <c>false</c>.
        /// </value>
        public bool Exists { get; private set; }

        /// <summary>
        /// Gets or sets the document.
        /// </summary>
        /// <value>
        /// The document.
        /// </value>
        public ProductionOrders Document { get; set; }

        /// <summary>
        /// Gets the lines.
        /// </summary>
        public ProductionOrderLines Lines
        {
            get
            {
                if (this.Document == null)
                {
                    return null;
                }

                return this._lines ?? (this._lines = new ProductionOrderLines(this.Document.Lines));
            }
        }

        #endregion Properties

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="ProductionOrderAdapter"/> class, with a blank (new, unsaved) document 
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        public ProductionOrderAdapter(Company company) : base(company)
        {
            this.Document = (ProductionOrders)this.Company.GetBusinessObject(BoObjectTypes.oProductionOrders);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ProductionOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="document">The document</param>
        public ProductionOrderAdapter(Company company, ProductionOrders document)
            : base(company)
        {
            this.Document = document;
            this.Exists = this.Document.DocumentNumber > 0;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ProductionOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="docEntry">The document id</param>
        public ProductionOrderAdapter(Company company, int docEntry) : base(company)
        {
            this.Document = (ProductionOrders)this.Company.GetBusinessObject(BoObjectTypes.oProductionOrders);
            this.Exists = this.Document.GetByKey(docEntry);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ProductionOrderAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index of the object to load from the XML.</param>
        public ProductionOrderAdapter(Company company, string xmlFile, int index) : base(company)
        {
            this.Document = (ProductionOrders) this.Company.GetBusinessObjectFromXML(xmlFile, index);
            this.Exists = this.Document.DocumentNumber > 0;
        }

        #endregion Constructor(s)

        #region Public Method(s)

        /// <summary>
        /// Adds this instance.
        /// </summary>
        /// <exception cref="SapException">Thrown if there is an SAP error adding the Document</exception>
        /// <returns>True of the method adds successfully</returns>
        public virtual bool Add()
        {
            this.Task = "Adding Production Order";
            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the ProductionOrder.Add is not set.");
            }

            if (this.Document == null)
            {
                throw new NullReferenceException("The Production Order has not been set. ProductionOrder.Add");
            }

            int counter = 0;
            SapException ex = null;
            while (counter < 5)
            {
                this.Document.SaveToFile(@"C:\tmp\samplePO.xml");
                if (this.Document.Add() == 0)
                {
                    ex = null;
                    break;
                }

                ex = new SapException(this.Company);

                if (!ex.DoRetry)
                {
                    break;
                }

                this.Task = string.Format("System busy, the Production Order cannot be added at this time. Attempt {0}", counter + 1);
                System.Threading.Thread.Sleep(1000);
                counter++;
            }

            if (ex != null)
            {
                throw ex;
            }

            string newDocEntry = this.Company.GetNewObjectKey();

            ThreadedAppLog.WriteLine("Production Order successfully created with doc entry {0}", newDocEntry);

            int docEntry = Convert.ToInt32(newDocEntry);
            this.Document = this.LoadDocument(docEntry);

            return true;
        }

        /// <summary>
        /// Cancels the document.
        /// </summary>
        /// <exception cref="SapException">Thrown if there is an SAP error closing the Document</exception>
        public virtual void Cancel()
        {
            this.Task = "Cancelling Production Order [" + this.Document.AbsoluteEntry + "]";
            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the ProductionOrderAdapter.Cancel is not set.");
            }

            if (this.Document == null)
            {
                throw new NullReferenceException("The Production Order has not been set. ProductionOrderAdapter.Cancel");
            }

            int counter = 0;
            SapException ex = null;
            while (counter < 5)
            {
                if (this.Document.Cancel() == 0)
                {
                    ex = null;
                    break;
                }

                ex = new SapException(this.Company);

                if (!ex.DoRetry)
                {
                    break;
                }

                this.Task = string.Format("System busy, the document cannot be closed at this time. Attempt {0}", counter + 1);
                System.Threading.Thread.Sleep(1000);
                counter++;
            }

            if (ex != null)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Updates this instance.
        /// </summary>
        /// <exception cref="SapException">Thrown if there is an SAP error updating the Document</exception>
        /// <returns>True of the method updates successfully</returns>
        public bool Update()
        {
            this.Task = "Updating Production Order [" + this.Document.AbsoluteEntry + "]";

            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the ProductionOrderAdapter.Update is not set.");
            }

            if (this.Document == null)
            {
                throw new NullReferenceException("The production order has not been set. ProductionOrderAdapter.Update");
            }


            int counter = 0;
            SapException ex = null;
            while (counter < 5)
            {
                if (this.Document.Update() == 0)
                {
                    ex = null;
                    break;
                }

                ex = new SapException(this.Company);

                if (!ex.DoRetry)
                {
                    break;
                }

                this.Task = string.Format("System busy, the document cannot be updated at this time. Attempt {0}", counter + 1);
                System.Threading.Thread.Sleep(1000);
                counter++;
            }

            if (ex != null)
            {
                throw ex;
            }

            return true;
        }

        /// <summary>
        /// Sets the user defined field.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <param name="fieldValue">The field value.</param>
        public void SetUserDefinedField(object fieldIndex, object fieldValue)
        {
            COMHelper.UserDefinedFieldValue(this.Document.UserFields, fieldIndex, fieldValue);
        }

        /// <summary>
        /// Sets the user defined field on the current line.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <param name="fieldValue">The field value.</param>
        public void SetLineUserDefinedField(object fieldIndex, object fieldValue)
        {
            ProductionOrders_Lines line = this.Document.Lines;
            try
            {
                UserFields uf = line.UserFields;
                try
                {
                    Fields fields = uf.Fields;
                    try
                    {
                        Field field = fields.Item(fieldIndex);
                        try
                        {
                            field.Value = fieldValue;
                        }
                        finally
                        {
                            COMHelper.Release(ref field);
                        }
                    }
                    finally
                    {
                        COMHelper.Release(ref fields);
                    }
                }
                finally
                {
                    COMHelper.Release(ref uf);
                }
            }
            finally
            {
                COMHelper.Release(ref line);
            }
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
        /// Gets the user defined field from the current line.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <returns>The value of the field</returns>
        public object GetLineUserDefinedField(object fieldIndex)
        {
            ProductionOrders_Lines line = this.Document.Lines;
            try
            {
                UserFields uf = line.UserFields;
                try
                {
                    Fields fields = uf.Fields;
                    try
                    {
                        Field field = fields.Item(fieldIndex);
                        try
                        {
                            return field.Value;
                        }
                        finally
                        {
                            COMHelper.Release(ref field);
                        }
                    }
                    finally
                    {
                        COMHelper.Release(ref fields);
                    }
                }
                finally
                {
                    COMHelper.Release(ref uf);
                }
            }
            finally
            {
                COMHelper.Release(ref line);
            }
        }

        #endregion Public Method(s)

        #region Protected Method(s)

        /// <summary>
        /// Loads a document from the SAP database
        /// </summary>
        /// <param name="docEntry">The document's ID (docEntry)</param>
        /// <returns>The document object</returns>
        protected virtual ProductionOrders LoadDocument(int docEntry)
        {
            // Pre-condition
            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the DocumentAdapter is not set.");
            }

            var document = (ProductionOrders)this.Company.GetBusinessObject(BoObjectTypes.oProductionOrders);
            this.Exists = document.GetByKey(docEntry);


            return document;
        }

        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected override void Release()
        {
            if (this._lines != null)
            {
                _lines.Dispose();
                _lines = null;
            }

            if (this.Document != null)
            {
                COMHelper.ReleaseOnly(this.Document);
                this.Document = null;
            }
        }

        #endregion Protected Method(s)
    }
}
