// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TableFieldAdapter.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   An object representing a SAP field
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using B1C.SAP.DI.BusinessAdapters.Administration;

namespace B1C.SAP.DI.BusinessAdapters
{
    using System;
    using Components;
    using Helpers;
    using SAPbobsCOM;
    using B1C.SAP.DI.Model.Exceptions;
using System.Collections.Generic;
    using System.Xml.Serialization;
    using System.IO;
    
    /// <summary>
    /// An object representing a SAP field
    /// </summary>
    public class TableFieldAdapter : SapObjectAdapter
    {
        /// <summary>
        /// True if the field already exists
        /// </summary>
        private bool exists;

        /// <summary>
        /// The field private
        /// </summary>
        private UserFieldsMD field;

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="TableFieldAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="fieldName">Name of the field.</param>
        public TableFieldAdapter(Company company, string tableName, string fieldName) : base(company)
        {
            this.Name = fieldName;
            this.TableName = tableName;
            this.ValidValues = new Dictionary<string, string>();
        }
        #endregion Constructors

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="TableFieldAdapter"/> is required.
        /// </summary>
        /// <value><c>true</c> if required; otherwise, <c>false</c>.</value>
        public bool Required{ get; set; }

        /// <summary>
        /// Gets or sets the id.
        /// </summary>
        /// <value>The field id.</value>
        public int Id { get; set; }

        /// <summary>
        /// Gets or sets the name of the table.
        /// </summary>
        /// <value>The name of the table.</value>
        public string TableName { get; set; }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>The field name.</value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the description.
        /// </summary>
        /// <value>The description.</value>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        /// <value>The field type.</value>
        public BoFieldTypes Type { get; set; }

        /// <summary>
        /// Gets or sets the type of the sub.
        /// </summary>
        /// <value>The type of the sub.</value>
        public BoFldSubTypes? SubType { get; set; }

        /// <summary>
        /// Gets or sets the size.
        /// </summary>
        /// <value>The field size.</value>
        public int? Size { get; set; }

        /// <summary>
        /// Gets or sets the valid values.
        /// </summary>
        /// <value>The valid values.</value>
        public Dictionary<string, string> ValidValues { get; set; }

        /// <summary>
        /// Gets or sets the default value.
        /// </summary>
        /// <value>The default value.</value>
        public string DefaultValue { get; set; }

        /// <summary>
        /// Gets or sets the size of the edit.
        /// </summary>
        /// <value>The size of the edit.</value>
        public int? EditSize { get; set; }

        /// <summary>
        /// Gets or sets the linked table.
        /// </summary>
        /// <value>The linked table.</value>
        public string LinkedTable { get; set; }

        /// <summary>
        /// Gets a value indicating whether this <see cref="TableFieldAdapter"/> is exists.
        /// </summary>
        /// <value><c>true</c> if exists; otherwise, <c>false</c>.</value>
        public bool Exists
        {
            get { return this.exists; }
        }

        /// <summary>
        /// Loads the specified table name.
        /// </summary>
        public void Load()
        {
            int fieldId = GetFieldId();

            this.field = (UserFieldsMD)this.Company.GetBusinessObject(BoObjectTypes.oUserFields);
            this.exists = this.field.GetByKey(this.TableName, fieldId);

            if (this.Exists)
            {
                this.Id = this.field.FieldID;
                this.Name = this.field.Name;
                this.Description = this.field.Description;
                this.TableName = this.field.TableName;
                this.Type = this.field.Type;
                this.SubType = this.field.SubType;
                this.Size = this.field.Size;
                this.EditSize = this.field.EditSize;
                this.LinkedTable = this.field.LinkedTable;
            }
        }

        /// <summary>
        /// Adds this instance.
        /// </summary>
        public void Add()
        {
            if (this.Name.Length > 20)
            {
                throw new Exception(string.Format("Field name {0} is too long. Max 20 characters allowed.", this.Name));
            }

            if (this.Name.Length == 0)
            {
                throw new Exception("Field name is not set..");
            }

            if (this.Description.Length > 30)
            {
                throw new Exception(string.Format("Field description {0} is too long. Max 30 characters allowed.", this.Description));
            }

            if (this.Description.Length == 0)
            {
                this.Description = this.Name;
            }

            if (this.field == null)
            {
                this.Load();

            }

            if (this.Exists)
            {
                return;
            }

            this.field.Name = this.Name;
            this.field.Description = this.Description;
            this.field.TableName = this.TableName;
            this.field.Type = this.Type;

            if (this.field.Type == BoFieldTypes.db_Alpha)
            {
                if (this.Size != null)
                {
                    this.field.Size = (int)this.Size;
                }

                if (this.EditSize != null)
                {
                    this.field.EditSize = (int)this.EditSize;
                }
            }

            if (this.SubType != null)
            {
                this.field.SubType = (BoFldSubTypes)this.SubType;
            }
            else
            {
                this.field.SubType = BoFldSubTypes.st_None;
            }

            if (!string.IsNullOrEmpty(this.LinkedTable))
            {
                this.field.LinkedTable = this.LinkedTable;
            }

            string defaultValue = null;
            if (this.ValidValues != null)
            {
                foreach (KeyValuePair<string, string> validValue in this.ValidValues)
                {
                    if (defaultValue == null)
                    {
                        defaultValue = validValue.Key;
                    }

                    field.ValidValues.Description = validValue.Value;
                    field.ValidValues.Value = validValue.Key;
                    field.ValidValues.Add();
                }
            }

            if (this.DefaultValue != null)
            {
                this.Required = true;
                this.field.DefaultValue = this.DefaultValue;
            }
            else
            {
                if (defaultValue != null)
                {
                    this.Required = true;
                    this.field.DefaultValue = defaultValue;
                    this.DefaultValue = defaultValue;
                }
            }

            this.field.Mandatory = this.Required ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;

            if (!this.Exists)
            {
                if (this.field.Add() != 0)
                {
#if DEBUG
                    string fileName = @"C:\temp\Invalid Fields\" + this.Name + ".Xml";
                    this.field.SaveXML(ref fileName);
                    var ex = new SapException(this.Company);
                    System.Diagnostics.Debug.WriteLine(ex.Message);
#else
                    throw new SapException(this.Company);

#endif 
                }
            }
        }

        /// <summary>
        /// Gets the field id.
        /// </summary>
        /// <returns>The id of the field</returns>
        public int GetFieldId()
        {
            int retValue;

            // Check if it is in a UDT
            using (var recordset = new RecordsetAdapter(this.Company, string.Format("SELECT * FROM CUFD WHERE TableId = '@{0}' AND AliasID = '{1}'", this.TableName, this.Name)))
            {
                if (recordset.EoF)
                {
                    // Check in the SAP Standard Tables
                    using (var recordsetStandard = new RecordsetAdapter(this.Company, string.Format("SELECT * FROM CUFD WHERE TableId = '{0}' AND AliasID = '{1}'", this.TableName, this.Name)))
                    {
                        if (recordsetStandard.EoF)
                        {
                            retValue = -1;
                        }
                        else
                        {
                            retValue = Convert.ToInt32(recordset.FieldValue("FieldID").ToString());
                        }
                    }
                }
                else
                {
                    retValue = Convert.ToInt32(recordset.FieldValue("FieldID").ToString());
                }
            }

            return retValue;
        }

        /// <summary>
        /// Gets the next code.
        /// </summary>
        /// <param name="prefix">The prefix.</param>
        /// <param name="format">The format.</param>
        /// <returns>The next available code.</returns>
        public string NextCode(string prefix, string format)
        {
            string query = string.Format(
                "SELECT ISNULL(MAX(CAST(REPLACE([{1}], '{2}', '') AS INT)),0) + 1 FROM [{0}] Where [{1}] like '{2}%'",
                this.TableName,
                this.Name,
                prefix);

            using (var rs = new RecordsetAdapter(this.Company, query))
            {
                int maxId = rs.EoF ? 1 : int.Parse(rs.FieldValue(0).ToString());
                return string.Format("{0}{1}", prefix, maxId.ToString("00000"));
            }
        }


        /// <summary>
        /// Gets the next code.
        /// </summary>
        /// <param name="prefix">The prefix.</param>
        /// <param name="format">The format.</param>
        /// <param name="plusValue">The plus value.</param>
        /// <returns>The next available code.</returns>
        public string NextCode(string prefix, string format, int plusValue)
        {
            string query = string.Format(
                "SELECT ISNULL(MAX(CAST(REPLACE([{1}], '{2}', '') AS INT)),0) + 1 FROM [{0}] Where [{1}] like '{2}%'",
                this.TableName,
                this.Name,
                prefix);

            using (var rs = new RecordsetAdapter(this.Company, query))
            {
                int maxId = rs.EoF ? 1 + plusValue : int.Parse(rs.FieldValue(0).ToString()) + plusValue;
                return string.Format("{0}{1}", prefix, maxId.ToString("00000"));
            }
        }

        /// <summary>
        /// Determines whether the specified primary key is translated.
        /// </summary>
        /// <param name="primaryKey">The primary key.</param>
        /// <param name="langCode">The lang code.</param>
        /// <param name="currentValue">The current value.</param>
        /// <returns>
        /// <c>true</c> if the specified primary key is translated; otherwise, <c>false</c>.
        /// </returns>
        public bool IsTranslated(string primaryKey, int langCode, ref object currentValue)
        {
            string query = "SELECT T1.Trans FROM [OMLT] T0 JOIN [MLT1] T1 ON T0.TranEntry = T1.TranEntry ";
            query += "WHERE T0.[TableName] = '{0}' AND T0.[FieldAlias] = '{1}' ";
            query += "AND T0.PK='{2}' AND T1.LangCode = {3}";

            query = string.Format(
                query,
                this.TableName,
                this.Name,
                primaryKey,
                langCode);

            using (var translationRecordset = new RecordsetAdapter(this.Company, query))
            {
                if (!translationRecordset.EoF)
                {
                    currentValue = translationRecordset.FieldValue("Trans").ToString();
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Gets the translation entry.
        /// </summary>
        /// <param name="primaryKey">The primary key.</param>
        /// <returns>The translation entry</returns>
        public string TranslationEntry(string primaryKey)
        {
            string query = "SELECT TranEntry FROM OMLT WHERE TableName = '{0}' AND FieldAlias = '{1}' AND PK = '{2}'";
            query = string.Format(query, this.TableName, this.Name, primaryKey);

            using (var translationRecordset = new RecordsetAdapter(this.Company, query))
            {
                if (!translationRecordset.EoF)
                {
                    return translationRecordset.FieldValue("TranEntry").ToString();
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Moves the translation.
        /// </summary>
        /// <param name="fromKey">Current From key.</param>
        /// <param name="toTable">Current To table.</param>
        /// <param name="toKey">Current To key.</param>
        /// <returns>True if moves the translation.</returns>
        public bool MoveTranslation(string fromKey, string toTable, string toKey)
        {
            return this.MoveTranslation(fromKey, toTable, this.Name, toKey);
        }

        /// <summary>
        /// Moves the translation.
        /// </summary>
        /// <param name="fromKey">Current From key.</param>
        /// <param name="toTable">Current To table.</param>
        /// <param name="toAlias">Current To alias.</param>
        /// <param name="toKey">Current To key.</param>
        /// <returns>True if moves the translation.</returns>
        public bool MoveTranslation(string fromKey, string toTable, string toAlias, string toKey)
        {
            this.Company.StartTransaction();

            string query = "SELECT * FROM OMLT WHERE TableName = '{0}' AND FieldAlias = '{1}' AND PK = '{2}'";
            query = string.Format(query, this.TableName, this.Name, fromKey);
            using (var translationRecordset = new RecordsetAdapter(this.Company, query))
            {
                string key = string.Empty;

                if (!translationRecordset.EoF)
                {
                    var fromTrans = (MultiLanguageTranslations)this.Company.GetBusinessObject(BoObjectTypes.oMultiLanguageTranslations);
                    if (fromTrans.GetByKey(Convert.ToInt32(translationRecordset.FieldValue("TranEntry").ToString())))
                    {
                        key = new TableFieldAdapter(this.Company, toTable, toAlias).TranslationEntry(toKey);
                        if (string.IsNullOrEmpty(key))
                        {
                            key = new TableFieldAdapter(this.Company, "OMLT", "TranEntry").NextCode(string.Empty, string.Empty);
                            string insertMlt = "INSERT INTO [OMLT]([TranEntry],[TableName],[FieldAlias],[PK],[DataSource],[UserSign]) VALUES({0},'{1}','{2}','{3}','I',{4})";
                            insertMlt = string.Format(insertMlt, Convert.ToInt32(key), toTable, toAlias, toKey, UserAdapter.Current(this.Company).InternalKey);
                            RecordsetAdapter.Execute(this.Company, insertMlt);
                            TableAdapter.SetDocEntry(this.Company, Convert.ToInt32(key) + 1, "224");
                        }

                        for (int transIndex = 0; transIndex < fromTrans.TranslationsInUserLanguages.Count; transIndex++)
                        {
                            fromTrans.TranslationsInUserLanguages.SetCurrentLine(transIndex);

                            // Check if translation exists
                            object existingTranslation = string.Empty;
                            if (!new TableFieldAdapter(this.Company, toTable, toAlias).IsTranslated(toKey, fromTrans.TranslationsInUserLanguages.LanguageCodeOfUserLanguage, ref existingTranslation))
                            {
                                string insertMlt1 = "INSERT INTO [MLT1] ([TranEntry],[LangCode],[Trans]) VALUES({0},{1},'{2}')";
                                insertMlt1 = string.Format(insertMlt1, Convert.ToInt32(key), fromTrans.TranslationsInUserLanguages.LanguageCodeOfUserLanguage, fromTrans.TranslationsInUserLanguages.Translationscontent.Replace("'", "''"));
                                RecordsetAdapter.Execute(this.Company, insertMlt1);
                            }
                        }
                    }
                }
            }

            this.Company.EndTransaction(BoWfTransOpt.wf_Commit);
            return true;
        }

        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected override void Release()
        {
            COMHelper.Release(ref this.field);
        }
    }
}
