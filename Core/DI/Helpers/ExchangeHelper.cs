//-----------------------------------------------------------------------
// <copyright file="ExchangeHelper.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using BusinessAdapters;
    using Components;
    using SAPbobsCOM;
    using Utility.Logging;
    using B1C.SAP.DI.Model.Exceptions;

    /// <summary>
    /// Instance of the ExchangeHelper class.
    /// </summary>
    /// <author>Frank Alunni</author>
    /// <remarks>This class has the function of retrieving and setting data and data objects in and out of SAP. Including transactions, recordsets, table definition, field definition and object registration definitions.</remarks>
    public class ExchangeHelper
    {
        #region Fields

        /// <summary>
        /// The SAP company
        /// </summary>
        private readonly Company company;

        /// <summary>
        /// Instance of the ExchangeHelper.
        /// </summary>
        private static ExchangeHelper instance;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeHelper"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        public ExchangeHelper(Company company)
        {
            this.company = company;
        }

        #endregion Constructors

        #region Methods
        /// <summary>
        /// Instances the ExchangeHelper class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns>An Instance to the ExchangeHelper class</returns>
        public static ExchangeHelper Instance(Company company)
        {
            if (instance == null)
            {
                instance = new ExchangeHelper(company);
            }

            return instance;
        }

        /// <summary>
        /// Creates the field.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="fieldDescription">The field description.</param>
        /// <param name="fieldType">Type of the field.</param>
        /// <param name="fieldSize">Size of the field.</param>
        /// <param name="fieldEditSize">Size of the field edit.</param>
        /// <param name="fieldSubType">Type of the field sub.</param>
        /// <param name="linkedTable">The linked table.</param>
        /// <returns>True if the field is created.</returns>
        [Obsolete("This method should be avoided. Use B1C.SAP.DI.BusinessAdapters.Administration.TableFieldAdapter.Add")]
        public bool CreateField(
            string tableName,
            string fieldName,
            string fieldDescription,
            BoFieldTypes fieldType,
            int? fieldSize,
            int? fieldEditSize,
            BoFldSubTypes? fieldSubType,
            string linkedTable)
        {
            // Try to retrieve if exists
            UserFieldsMD field = this.GetField(tableName, fieldName);

            try
            {
                if (field == null)
                {
                    // create new table
                    field = (UserFieldsMD)this.company.GetBusinessObject(BoObjectTypes.oUserFields);
                    field.TableName = tableName;

                    field.Name = fieldName.Length > 20 ? fieldName.Substring(0, 20) : fieldName;

                    if (fieldDescription.Length > 30)
                    {
                        fieldDescription = fieldDescription.Substring(0, 27) + "...";
                    }

                    field.Description = fieldDescription;
                    field.Type = fieldType;
                    if (fieldSize != null)
                    {
                        field.Size = (int)fieldSize;
                    }

                    if (fieldEditSize != null)
                    {
                        field.EditSize = (int)fieldEditSize;
                    }

                    if (fieldSubType != null)
                    {
                        field.SubType = (BoFldSubTypes)fieldSubType;
                    }
                    else
                    {
                        field.SubType = BoFldSubTypes.st_None;
                    }

                    if (!string.IsNullOrEmpty(linkedTable))
                    {
                        field.LinkedTable = linkedTable;
                    }

                    if (field.Add() != 0)
                    {
                        throw new SapException(this.company);
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return true;
                }
            }
            finally
            {
                if (field != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(field);
                }

                GC.Collect();
            }
        }

        /// <summary>
        /// Creates the field.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="fieldDescription">The field description.</param>
        /// <param name="fieldType">Type of the field.</param>
        /// <param name="fieldSize">Size of the field.</param>
        /// <param name="fieldEditSize">Size of the field edit.</param>
        /// <param name="fieldSubType">Type of the field sub.</param>
        /// <param name="linkedTable">The linked table.</param>
        /// <param name="validValues">The valid values.</param>
        /// <returns>True if the field is created.</returns>
        [Obsolete("This method should be avoided. Use B1C.SAP.DI.BusinessAdapters.Administration.TableFieldAdapter.Add")]
        public bool CreateField(
            string tableName,
            string fieldName,
            string fieldDescription,
            BoFieldTypes fieldType,
            int? fieldSize,
            int? fieldEditSize,
            BoFldSubTypes? fieldSubType,
            string linkedTable,
            Dictionary<string, string> validValues)
        {
            // Try to retrieve if exists
            UserFieldsMD field = this.GetField(tableName, fieldName);

            try
            {
                if (field == null)
                {
                    // create new table
                    field = (UserFieldsMD)this.company.GetBusinessObject(BoObjectTypes.oUserFields);
                    field.TableName = tableName;

                    field.Name = fieldName.Length > 20 ? fieldName.Substring(0, 20) : fieldName;

                    if (fieldDescription.Length > 30)
                    {
                        fieldDescription = fieldDescription.Substring(0, 27) + "...";
                    }

                    field.Description = fieldDescription;
                    field.Type = fieldType;
                    if (fieldSize != null)
                    {
                        field.Size = (int)fieldSize;
                    }

                    if (fieldEditSize != null)
                    {
                        field.EditSize = (int)fieldEditSize;
                    }

                    if (fieldSubType != null)
                    {
                        field.SubType = (BoFldSubTypes)fieldSubType;
                    }
                    else
                    {
                        field.SubType = BoFldSubTypes.st_None;
                    }

                    if (!string.IsNullOrEmpty(linkedTable))
                    {
                        field.LinkedTable = linkedTable;
                    }

                    foreach (KeyValuePair<string, string> validValue in validValues)
                    {
                        field.ValidValues.Description = validValue.Value;
                        field.ValidValues.Value = validValue.Key;
                        field.ValidValues.Add();
                    }

                    if (field.Add() != 0)
                    {
                        throw new SapException(this.company);
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return true;
                }
            }
            finally
            {
                if (field != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(field);
                }

                GC.Collect();
            }
        }

        /// <summary>
        /// Creates the table.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="tableDescription">The table description.</param>
        /// <param name="tableType">Type of the table.</param>
        /// <returns>True is the table is created.</returns>
        [Obsolete("This method should be avoided. Use B1C.SAP.DI.BusinessAdapters.Administration.TableAdapter.Add")]
        public bool CreateTable(string tableName, string tableDescription, BoUTBTableType tableType)
        {
            // Try to retrieve if exists
            UserTablesMD table = this.GetTable(tableName);

            try
            {
                GC.Collect();
                if (table == null)
                {
                    // create new table
                    table = (UserTablesMD)this.company.GetBusinessObject(BoObjectTypes.oUserTables);
                    if (tableName.Length > 19)
                    {
                        tableName = tableName.Substring(0, 19);
                    }

                    table.TableName = tableName;
                    if (tableDescription.Length > 30)
                    {
                        tableDescription = tableDescription.Substring(0, 27) + "...";
                    }

                    table.TableDescription = tableDescription;
                    table.TableType = tableType;
                    
                    if (table.Add() != 0)
                    {
                        throw new SapException(this.company);
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }
            finally
            {
                if (table != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                    GC.Collect();
                }
            }
        }

        /// <summary>
        /// Gets the next code.
        /// </summary>
        /// <param name="objectName">The name of the object</param>
        /// <param name="prefix">The prefix.</param>        
        /// <param name="numberLength">The length of the code string number part</param>
        /// <param name="objectCode">The code for the next object</param>
        /// <param name="docEntry">The next docEntry</param>
        public void GetNextCode(string objectName, string prefix, int numberLength, out string objectCode, out int docEntry)
        {
            var stringFormat = new StringBuilder();

            // Build out a string like "000000" to use to format an int
            for (int i = 0; i < numberLength; i++)
            {
                stringFormat.Append("0");
            }

            int maxId = TableAdapter.NextObjectKey(this.company, objectName);
            docEntry = maxId;

            objectCode = numberLength > 0 ? string.Format("{0}{1}", prefix, maxId.ToString(stringFormat.ToString())) : maxId.ToString();           
        }

        /// <summary>
        /// Gets the table.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <returns>The table selected.</returns>
        public UserTablesMD GetTable(string tableName)
        {
            if (this.company == null)
            {
                return null;
            }

            // Try to retrieve if exists
            var table = (UserTablesMD)this.company.GetBusinessObject(BoObjectTypes.oUserTables);
            try
            {
                if (!table.GetByKey(tableName))
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                    table = null;
                    GC.Collect();

                    // If the table cannot be return than return null
                    return null;
                }

                return table;
            }
            catch
            {
                if (table != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                }

                GC.Collect();

                return null;
            }
        }

        /// <summary>
        /// Registers the table.
        /// </summary>
        /// <param name="objectName">Name of the object.</param>
        /// <param name="objectCode">The object code.</param>
        /// <param name="objectType">Type of the object.</param>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="childTablesCsv">The child tables CSV.</param>
        /// <param name="canDelete">if set to <c>true</c> [can delete].</param>
        /// <param name="canCancel">if set to <c>true</c> [can cancel].</param>
        /// <param name="canClose">if set to <c>true</c> [can close].</param>
        /// <param name="canCreateDefaultForm">if set to <c>true</c> [can create default form].</param>
        /// <param name="canFind">if set to <c>true</c> [can find].</param>
        /// <param name="canYearTransfer">if set to <c>true</c> [can year transfer].</param>
        /// <param name="manageSeries">if set to <c>true</c> [manage series].</param>
        /// <param name="findColumnsCsv">The find columns CSV.</param>
        /// <param name="formColumnsCsv">The form columns CSV.</param>
        /// <returns>True if the table is registered</returns>
        public bool RegisterTable(
            string objectName,
            string objectCode,
            BoUDOObjType objectType,
            string tableName,
            string childTablesCsv,
            bool canDelete,
            bool canCancel,
            bool canClose,
            bool canCreateDefaultForm,
            bool canFind,
            bool canYearTransfer,
            bool manageSeries,
            string findColumnsCsv,
            string formColumnsCsv)
        {
            using (var recordset = new RecordsetAdapter(this.company, string.Format("SELECT * FROM OUDO WHERE Code = '{0}'", objectCode)))
            {
                if (!recordset.EoF)
                {
                    return false;
                }
            }

            var reg = (UserObjectsMD)this.company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

            try
            {
                reg.CanCancel = this.SAPBoolean(canCancel);
                reg.CanClose = this.SAPBoolean(canClose);
                reg.CanDelete = this.SAPBoolean(canDelete);
                reg.CanYearTransfer = this.SAPBoolean(canYearTransfer);
                reg.ManageSeries = this.SAPBoolean(manageSeries);
                reg.Code = objectCode;
                reg.Name = objectName;
                reg.ObjectType = objectType;
                reg.TableName = tableName;

                if (!string.IsNullOrEmpty(childTablesCsv))
                {
                    string[] childTables = childTablesCsv.Split(',');
                    for (int childIndex = 0; childIndex < childTables.Length; childIndex++)
                    {
                        if (childIndex >= reg.ChildTables.Count)
                        {
                            reg.ChildTables.Add();
                        }

                        reg.ChildTables.SetCurrentLine(childIndex);
                        reg.ChildTables.TableName = childTables[childIndex];
                    }
                }

                // Add the find columns
                if (!string.IsNullOrEmpty(findColumnsCsv))
                {
                    string[] findColumns = findColumnsCsv.Split(',');
                    for (int colIndex = 0; colIndex < findColumns.Length; colIndex++)
                    {
                        if (colIndex >= reg.FindColumns.Count)
                        {
                            reg.FindColumns.Add();
                        }

                        reg.FindColumns.SetCurrentLine(colIndex);
                        reg.FindColumns.ColumnAlias = findColumns[colIndex];
                    }

                    reg.CanFind = this.SAPBoolean(canFind);
                }

                // Add the form columns
                if (!string.IsNullOrEmpty(formColumnsCsv))
                {
                    string[] formColumns = formColumnsCsv.Split(',');
                    for (int colIndex = 0; colIndex < formColumns.Length; colIndex++)
                    {
                        if (colIndex >= reg.FormColumns.Count)
                        {
                            reg.FormColumns.Add();
                        }

                        reg.FormColumns.SetCurrentLine(colIndex);
                        reg.FormColumns.FormColumnAlias = formColumns[colIndex];
                    }

                    reg.CanCreateDefaultForm = this.SAPBoolean(canCreateDefaultForm);
                }

                if (reg.Add() != 0)
                {
                    // ErrorHandlerHelper.HandleBusinessError(this.company);
                    return false;
                }

                return true;
            }
            catch
            {
                // Do not catch errors some users cannot register objects
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(reg);
                GC.Collect();
            }
        }

        /// <summary>
        /// SAPs the boolean.
        /// </summary>
        /// <param name="inputValue">if set to <c>true</c> [input value].</param>
        /// <returns>The SAPbobsCOM BoYesNoEnum.</returns>
        public BoYesNoEnum SAPBoolean(bool inputValue)
        {
            if (inputValue)
            {
                return BoYesNoEnum.tYES;
            }

            return BoYesNoEnum.tNO;
        }

        /// <summary>
        /// Sends the email.
        /// </summary>
        /// <param name="emailSubject">The email subject.</param>
        /// <param name="emailBody">The email body.</param>
        /// <param name="internalRecipients">The internal recipients.</param>
        /// <param name="externalRecipients">The external recipients.</param>
        /// <param name="priority">The priority.</param>
        public void SendEmail(string emailSubject, string emailBody, string internalRecipients, string externalRecipients, BoMsgPriorities priority)
        {
            try
            {
                var message = (Messages)this.company.GetBusinessObject(BoObjectTypes.oMessages);
                const string External_user_code = "listener";

                message.Subject = emailSubject;
                message.MessageText = emailBody;

                // Add each recipient in the comma separated list
                string[] internalRecipientList = internalRecipients.Split(',');
                int recipientCount = 0;

                foreach (string recipient in internalRecipientList)
                {
                    string trimmedRecipient = recipient.Trim();
                    if (!string.IsNullOrEmpty(trimmedRecipient))
                    {
                        if (recipientCount > 0)
                        {
                            message.Recipients.Add();
                            message.Recipients.SetCurrentLine(recipientCount);
                        }

                        message.Recipients.UserCode = trimmedRecipient;
                        message.Recipients.SendInternal = BoYesNoEnum.tYES;
                        recipientCount++;
                    }
                }

                string[] externalRecipientList = externalRecipients.Split(',');

                foreach (string recipient in externalRecipientList)
                {
                    string trimmedRecipient = recipient.Trim();
                    if (!string.IsNullOrEmpty(trimmedRecipient))
                    {
                        if (recipientCount > 0)
                        {
                            message.Recipients.Add();
                            message.Recipients.SetCurrentLine(recipientCount);
                        }

                        message.Recipients.UserCode = External_user_code;
                        message.Recipients.EmailAddress = trimmedRecipient;
                        message.Recipients.SendEmail = BoYesNoEnum.tYES;
                        recipientCount++;
                    }
                }

                message.Priority = priority;

                if (message.Add() != 0)
                {
                    int errorId;
                    string errorMessage;
                    this.company.GetLastError(out errorId, out errorMessage);
                    var ex = new Exception(string.Format("SAP ERROR: [Code:{0}] [Message:{1}]", errorId, errorMessage));
                    ThreadedAppLog.WriteLine("Could not send email.");
                    ThreadedAppLog.WriteErrorLine(ex);
                }

                return;
            }
            catch (Exception ex)
            {
                // If an error occurr, log it and fails the method
                ThreadedAppLog.WriteLine("Could not send email.");
                ThreadedAppLog.WriteErrorLine(ex);
                return;
            }
        }

        /// <summary>
        /// Gets the field.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>The field requested if exists</returns>
        private UserFieldsMD GetField(string tableName, string fieldName)
        {
            if (this.company == null)
            {
                return null;
            }

            // Try to retrieve if exists
            var field = (UserFieldsMD)this.company.GetBusinessObject(BoObjectTypes.oUserFields);

            try
            {
                int fieldId = new TableFieldAdapter(this.company, tableName, fieldName).GetFieldId();

                if (fieldId < 0)
                {
                    return null;
                }

                if (!field.GetByKey(tableName, fieldId))
                {
                    // If the table cannot be return than return null
                    return null;
                }

                return field;
            }
            catch
            {
                return null;
            }
            finally
            {
                COMHelper.Release(ref field);
                GC.Collect();
            }
        }
        #endregion Methods
    }
}