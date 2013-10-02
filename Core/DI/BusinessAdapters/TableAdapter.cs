// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TableAdapter.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   An object representing a SAP table
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using B1C.Utility.Helpers;

namespace B1C.SAP.DI.BusinessAdapters
{
    using System;
    using Components;
    using Helpers;
    using SAPbobsCOM;
    using B1C.SAP.DI.Model.Exceptions;
    using B1C.SAP.DI.Model.EventArguments;
    using System.Runtime.InteropServices;

    /// <summary>
    /// An object representing a SAP table
    /// </summary>
    public class TableAdapter : SapObjectAdapter 
    {
        #region Fields
        /// <summary>
        /// The table private field
        /// </summary>
        private UserTablesMD _table;

        /// <summary>
        /// The fields collection private field
        /// </summary>
        private TableFieldAdapters _fields;
        #endregion

        #region Delegates
        /// <summary>
        /// The handler for the statusbar changed
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The StatusbarChangedEventArgs arguments.</param>
        public delegate void StatusbarChangedEventHandler(object sender, StatusbarChangedEventArgs e);
        #endregion Delegates

        #region Events
        /// <summary>
        /// Occurs when [statusbar change].
        /// </summary>
        public event StatusbarChangedEventHandler StatusbarChanged;
        #endregion Events

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="TableAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="tableName">Name of the table.</param>
        public TableAdapter(Company company, string tableName)
            : base(company)
        {
            this.Name = tableName;
            this.Registration = new TableRegistration(company, tableName);
        }
        #endregion Constructors

        #region Properties
        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>The table name.</value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the description.
        /// </summary>
        /// <value>The description.</value>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        /// <value>The table type.</value>
        public BoUTBTableType Type { get; set; }

        /// <summary>
        /// Gets a value indicating whether this <see cref="TableAdapter"/> is exists.
        /// </summary>
        /// <value><c>true</c> if exists; otherwise, <c>false</c>.</value>
        public bool Exists { get; private set; }

        /// <summary>
        /// Gets or sets the registration.
        /// </summary>
        /// <value>The registration.</value>
        public TableRegistration Registration { get; private set; }

        /// <summary>
        /// Gets or sets the fields.
        /// </summary>
        /// <value>The fields.</value>
        public TableFieldAdapters Fields
        {
            get
            {
                if (this._fields == null)
                {
                    this._fields = new TableFieldAdapters();
                }

                return this._fields;
            }

            set
            {
                this._fields = value;
            }
        }
        #endregion

        #region Methods

        /// <summary>
        /// Raises the StatusbarChanged event.
        /// </summary>
        /// <param name="e">The <see cref="Maritz.SAP.DI.Components.StatusbarChangedEventArgs"/> instance containing the event data.</param>
        protected virtual void OnStatusbarChanged(StatusbarChangedEventArgs e)
        {
            this.StatusbarChanged(this, e);
        }

        /// <summary>
        /// Gets the next available object key for the given object code
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="objectCode">The code of the type of object to retrieve the next key for</param>
        /// <returns>The next object key</returns>
        public static int NextObjectKey(Company company, string objectCode)
        {
            int toReturn = -1;
            using (var recordset = new RecordsetAdapter(company, string.Format("exec dbo.oz_GetNextObjectKey '{0}'", objectCode)))
            {
                if (!recordset.EoF)
                {
                    toReturn = int.Parse(recordset.FieldValue(0).ToString());
                }
            }

            return toReturn;
        }

        /// <summary>
        /// Sets the doc entry.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="docEntry">The doc entry.</param>
        /// <param name="objectName">Name of the object.</param>
        public static void SetDocEntry(Company company, int docEntry, string objectName)
        {
            using (var docEntryRecordset = new RecordsetAdapter(company))
            {
                docEntryRecordset.Execute(string.Format("UPDATE T0 SET T0.[AutoKey] = {0} FROM [dbo].[ONNM] T0  WHERE T0.[ObjectCode] = '{1}'", docEntry, objectName));
            }
        }

        /// <summary>
        /// Synchronizes the next key.
        /// </summary>
        /// <param name="objectCode">The object code.</param>
        /// <param name="keyColumn">The key column.</param>
        public void SynchronizeNextKey(string objectCode, string keyColumn)
        {
            var helper = new StoredProcHelper("dbo.oz_SynchAutoKey");
            helper.AddParameter("@tableName", System.Data.DbType.String, this.Name);
            helper.AddParameter("@objectCode", System.Data.DbType.String, objectCode);
            helper.AddParameter("@keyColumn", System.Data.DbType.String, keyColumn);

            RecordsetAdapter.Execute(this.Company, helper.ToString());
        }

        /// <summary>
        /// Gets the next line id.
        /// </summary>
        /// <param name="code">The object code.</param>
        /// <returns>The next line Id</returns>
        public int NextLineId(string code)
        {
            string query = string.Format(
                "SELECT ISNULL(MAX(LineId),0) + 1 FROM [{0}] Where Code ='{1}'",
                this.Name,
                code);

            using (var rs = new RecordsetAdapter(this.Company, query))
            {
                return rs.EoF ? 1 : int.Parse(rs.FieldValue(0).ToString());
            }
        }

        /// <summary>
        /// Loads this instance.
        /// </summary>
        public void Load()
        {
            this._table = (UserTablesMD)this.Company.GetBusinessObject(BoObjectTypes.oUserTables);
            this.Exists = this._table.GetByKey(this.Name);

            if (this.Exists)
            {
                this.Description = this._table.TableDescription;
                this.Name = this._table.TableName;
                this.Type = this._table.TableType;
            }
            else
            {
                // may not be a user table
                SAPbobsCOM.Recordset rs = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                try
                {
                    string query = string.Format("select count(1) from {0}", this.Name);
                    rs.DoQuery(query);
                    this.Exists = true;
                }
                catch
                {
                    // Ingore error table does not exists
                }
                finally
                {
                    Marshal.ReleaseComObject(rs);
                    rs = null;
                    GC.Collect();
                }
            }

            switch (this.Type)
            {
                case BoUTBTableType.bott_MasterData:
                    this.Registration.ObjectType = BoUDOObjType.boud_MasterData;
                    break;
                case BoUTBTableType.bott_Document:
                    this.Registration.ObjectType = BoUDOObjType.boud_Document;
                    break;
            }
        }

        /// <summary>
        /// Adds this instance.
        /// </summary>
        public void Add()
        {
            if (this.Name == null)
            {
                this.Name = string.Empty;
            }

            if (this.Name.Length > 20)
            {
                throw new Exception(string.Format("Table name {0} is too long. Max 20 characters allowed.", this.Name));
            }

            if (this.Name.Length == 0)
            {
                throw new Exception("Table name is not set..");
            }

            if (this.Description == null)
            {
                this.Description = string.Empty;
            }

            if (this.Description.Length > 30)
            {
                throw new Exception(string.Format("Table description {0} is too long. Max 30 characters allowed.", this.Description));
            }

            if (this.Description.Length == 0)
            {
                this.Description = this.Name;
            }

            this.Load();

            if (this.Exists)
            {
                return;
            }

            this._table.TableDescription = this.Description;
            this._table.TableName = this.Name;
            this._table.TableType = this.Type;

            if (this.StatusbarChanged != null)
            {
                this.OnStatusbarChanged(new StatusbarChangedEventArgs(string.Format("Creating '{0}'.", this.Description), false));
            }


            // Create new table
            if (this._table.Add() != 0)
            {
                throw new SapException(this.Company);
            }

            COMHelper.Release(ref this._table);

            // Add all the fields;
            foreach (TableFieldAdapter field in this.Fields)
            {
                if (this.StatusbarChanged != null)
                {
                    this.OnStatusbarChanged(new StatusbarChangedEventArgs(string.Format("Adding '{0}' to '{1}'.", field.Description, this.Description), false));
                }

                System.Diagnostics.Debug.WriteLine("Creating Field: " + field.Name);
                field.Add();
                System.Diagnostics.Debug.WriteLine("Field: " + field.Name + " created.");
            }

            this.Fields.Dispose();
            this.Fields = null;

            if (string.IsNullOrEmpty(this.Registration.ObjectName))
            {
                this.Registration.ObjectName = this.Description;
            }

            if (string.IsNullOrEmpty(this.Registration.ObjectCode))
            {
                this.Registration.ObjectCode = this.Registration.ObjectName.Replace(" ", string.Empty);
            }

            if (this.StatusbarChanged != null)
            {
                this.OnStatusbarChanged(new StatusbarChangedEventArgs(string.Format("Registering '{0}'.", this.Registration.ObjectName), false));
            }
            this.Registration.Add();
            if (this.StatusbarChanged != null)
            {
                this.OnStatusbarChanged(new StatusbarChangedEventArgs(string.Format("{0} created successfully.", this.Description), false));
            }

            this.Registration.Dispose();
        }


        /// <summary>
        /// Update this instance.
        /// </summary>
        public void Update()
        {
            this.Load();

            if (!this.Exists)
            {
                return;
            }

            if (this.StatusbarChanged != null)
            {
                this.OnStatusbarChanged(new StatusbarChangedEventArgs(string.Format("Updating '{0}'.", this.Name), false));
            }

            COMHelper.Release(ref this._table);

            // Add all the fields;
            foreach (TableFieldAdapter field in this.Fields)
            {
                if (this.StatusbarChanged != null)
                {
                    this.OnStatusbarChanged(new StatusbarChangedEventArgs(string.Format("Adding '{0}' to '{1}'.", field.Description, this.Name), false));
                }

                System.Diagnostics.Debug.WriteLine("Creating Field: " + field.Name);
                field.Add();
                System.Diagnostics.Debug.WriteLine("Field: " + field.Name + " created.");
            }

            this.Fields.Dispose();
            this.Fields = null;
        }


        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected override void Release()
        {
            this.Fields.Dispose();
            this.Fields = null;

            this.Registration.Dispose();
            this.Registration = null;

            COMHelper.Release(ref this._table);
        }
        #endregion
    }
}
