//-----------------------------------------------------------------------
// <copyright file="RecordsetAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------

using B1C.SAP.DI.Helpers;

namespace B1C.SAP.DI.BusinessAdapters
{
    #region Using Directive(s)
    using SAPbobsCOM;
    using Utility.Helpers;
    using System;
    using B1C.Utility.Logging;

    #endregion Using Directive(s)

    /// <summary>
    /// The Recordset Adapter
    /// </summary>
    public class RecordsetAdapter : SapObjectAdapter
    {
        /// <summary>
        /// The COM Recordset object
        /// </summary>
        private Recordset recordset;

        /// <summary>
        /// Initializes a new instance of the <see cref="RecordsetAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="recordset">The recordset.</param>
        public RecordsetAdapter(Company company, Recordset recordset)
            : base(company)
        {
            this.Recordset = recordset;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RecordsetAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="query">The query.</param>
        public RecordsetAdapter(Company company, string query)
            : base(company)
        {
            try
            {
                this.Recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                this.Recordset.DoQuery(query);
            }
            catch (Exception ex)
            {
                ThreadedAppLog.WriteLine("Exception encountered running query: {0}", query);
                throw ex;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RecordsetAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        public RecordsetAdapter(Company company) : base(company)
        {
        }

        /// <summary>
        /// Gets a value indicating whether end of file.
        /// </summary>
        /// <value><c>true</c> if EOF; otherwise, <c>false</c>.</value>
        public bool EoF
        {
            get
            {
                if (this.Recordset == null)
                {
                    return true;
                }

                return this.Recordset.EoF;
            }
        }

        /// <summary>
        /// Gets a value indicating whether Beginning of File.
        /// </summary>
        /// <value><c>true</c> if BOF; otherwise, <c>false</c>.</value>
        public bool BoF
        {
            get
            {
                if (this.Recordset == null)
                {
                    return true;
                }

                return this.Recordset.BoF;
            }
        }

        /// <summary>
        /// Gets the record count.
        /// </summary>
        /// <value>The record count.</value>
        public int RecordCount
        {
            get
            {
                if (this.Recordset == null)
                {
                    return -1;
                }

                return this.Recordset.RecordCount;
            }
        }

        /// <summary>
        /// Gets or sets the recordset.
        /// </summary>
        /// <value>The recordset.</value>
        private Recordset Recordset
        {
            get
            {
                return this.recordset;
            }

            set
            {
                this.recordset = value;
            }
        }

        /// <summary>
        /// Executes the specified query.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="query">The query.</param>
        /// <returns>True if executes the specified query.</returns>
        public static bool Execute(Company company, string query)
        {
            if (company == null)
            {
                return false;
            }

            var rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                rs.DoQuery(query);
                return true;
            }
            catch (Exception ex)
            {
                ThreadedAppLog.WriteLine("Exception encountered running query: {0}", query);
                throw ex;
            }
            finally
            {
                COMHelper.Release(ref rs);
            }
        }

        /// <summary>
        /// Moves to the first record.
        /// </summary>
        public void MoveFirst()
        {
            if (this.Recordset == null)
            {
                return;
            }

            this.Recordset.MoveFirst();            
        }

        /// <summary>
        /// Moves to the last record.
        /// </summary>
        public void MoveLast()
        {
            if (this.Recordset == null)
            {
                return;
            }

            this.Recordset.MoveLast();
        }

        /// <summary>
        /// Moves to the next record.
        /// </summary>
        public void MoveNext()
        {
            if (this.Recordset == null)
            {
                return;
            }

            this.Recordset.MoveNext();
        }

        /// <summary>
        /// Moves to the previous record.
        /// </summary>
        public void MovePrevious()
        {
            if (this.Recordset == null)
            {
                return;
            }

            this.Recordset.MovePrevious();
        }

        /// <summary>
        /// Gets the field value.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <returns>The value of the field with the given index</returns>
        public object FieldValue(object fieldIndex)
        {
            if (this.Recordset == null)
            {
                return null;
            }

            Fields fields = this.Recordset.Fields;

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

        /// <summary>
        /// Executes the specified query.
        /// </summary>
        /// <param name="query">The query.</param>
        /// <returns>True if executes the specified query.</returns>
        public bool Execute(string query)
        {
            if (this.Company == null)
            {
                return false;
            }

            var rs = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                rs.DoQuery(query);
                return true;
            }
            catch (Exception ex)
            {            
                ThreadedAppLog.WriteLine("Exception encountered running query: {0}", query);
                throw ex;
            }
            finally
            {
                COMHelper.Release(ref rs);   
            }
        }

        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected override void Release()
        {
            COMHelper.Release(ref this.recordset);
        }
    }
}
