// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ActivityAdapter.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   The Activity Adapter
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.BusinessPartners
{
    using System;
    using System.Text;
    using Components;
    using Helpers;
    using SAPbobsCOM;
    using Utility.Helpers;
    using B1C.SAP.DI.Model.Exceptions;

    /// <summary>
    /// The Activity Adapter
    /// </summary>
    public class ActivityAdapter : SapObjectAdapter
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ActivityAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP company object</param>
        public ActivityAdapter(Company company) : base(company)
        {
            this.Document = (Contacts)this.Company.GetBusinessObject(BoObjectTypes.oContacts);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ActivityAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object.</param>
        /// <param name="activityCode">The activity code.</param>
        public ActivityAdapter(Company company, int activityCode) : this(company)
        {
            this.Load(activityCode);
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets or sets the document.
        /// </summary>
        /// <value>The document.</value>
        public Contacts Document { get; protected set; }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Instantiates a Activity Adapter with the given company
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <returns>An instance of the BusinessPartnerAdapter class</returns>
        public static ActivityAdapter Instance(Company company)
        {
            return new ActivityAdapter(company);
        }

        /// <summary>
        /// Instances the specified activity.
        /// </summary>
        /// <param name="company">The SAP company object.</param>
        /// <param name="activityCode">The activity code.</param>
        /// <returns>An instance to the BusinessPartnerAdapter class</returns>
        public static ActivityAdapter Instance(Company company, int activityCode)
        {
            return new ActivityAdapter(company, activityCode);
        }

        /// <summary>
        /// Loads the specified activity code.
        /// </summary>
        /// <param name="activityCode">The activity code.</param>
        /// <returns>True if the method loads successfully</returns>
        public bool Load(int activityCode)
        {
            if (this.Company == null)
            {
                throw new ArgumentNullException("activityCode");
            }

            return this.Document.GetByKey(activityCode);
        }

        /// <summary>
        /// Adds this instance.
        /// </summary>
        /// <returns>True if the method adds successfully</returns>
        public bool Add()
        {
            this.Task = "Adding new activity";
            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the ActivityAdapter.Add is not set.");
            }

            if (this.Document == null)
            {
                throw new NullReferenceException("The document has not been set. ActivityAdapter.Add");
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
            this.Task = "Updating activity";

            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the ActivityAdapter.Update is not set.");
            }

            if (this.Document == null)
            {
                throw new NullReferenceException("The document has not been set. ActivityAdapter.Update");
            }

            if (this.Document.Update() != 0)
            {
                throw new SapException(this.Company);
            }

            return true;
        }

        /// <summary>
        /// Creates the followup.
        /// </summary>
        /// <returns>The newly created document.</returns>
        public ActivityAdapter CreateFollowup()
        {
            this.Task = "Creating followup to activity";

            using (var newDocument = new ActivityAdapter(this.Company))
            {
                newDocument.Document.CardCode = this.Document.CardCode;
                newDocument.Document.Activity = this.Document.Activity;
                newDocument.Document.City = this.Document.City;
                newDocument.Document.ContactDate = this.Document.ContactDate;
                newDocument.Document.ContactPersonCode = this.Document.ContactPersonCode;
                newDocument.Document.ContactTime = this.Document.ContactTime;
                newDocument.Document.Country = this.Document.Country;
                string details = string.Format("Follow up to {0} -  {1}", this.Document.ContactCode, this.Document.Details);
                if (details.Length > 60)
                {
                    details = details.Substring(0, 57) + "...";
                }

                newDocument.Document.Duration = this.Document.Duration;
                newDocument.Document.DurationType = this.Document.DurationType;
                newDocument.Document.EndDuedate = DateTime.Today;
                newDocument.Document.Fax = this.Document.Fax;
                newDocument.Document.Notes = this.Document.Notes;
                newDocument.Document.Phone = this.Document.Phone;
                newDocument.Document.PreviousActivity = Convert.ToInt32(this.Document.ContactCode);
                newDocument.Document.Recontact = DateTime.Today;
                newDocument.Document.Priority = this.Document.Priority;
                newDocument.Document.Reminder = this.Document.Reminder;
                newDocument.Document.ReminderPeriod = this.Document.ReminderPeriod;
                newDocument.Document.ReminderType = this.Document.ReminderType;
                newDocument.Document.Room = this.Document.Room;
                newDocument.Document.StartDate = DateTime.Today;
                newDocument.Document.State = this.Document.State;
                newDocument.Document.Street = this.Document.Street;
                newDocument.Document.Details = details;
                newDocument.Document.Duration = this.Document.Duration;

                // Transfer all the custom fields value
                this.TransferFieldsToAdapter(newDocument);

                if (newDocument.Document.Add() != 0)
                {
                    throw new SapException(this.Company);
                }
            }

            string newDocEntry = this.Company.GetNewObjectKey();

            this.Task = "Linking followup activity to current document";

            // Add all the link documents
            var sql = new SqlHelper();
            sql.Builder.AppendFormat("SELECT * FROM [@XX_DOCACTIVITIES] WHERE [U_ActEntry] = '{0}'", this.Document.ContactCode);
            using (var rs = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                while (!rs.EoF)
                {
                    this.AddMultipleActivity(
                        Convert.ToInt32(newDocEntry),
                        Convert.ToInt32(rs.FieldValue("U_DocEntry").ToString()),
                        rs.FieldValue("U_ObjType").ToString(),
                        Convert.ToInt32(rs.FieldValue("UserSign").ToString()));
                    rs.MoveNext();
                }
            }

            return new ActivityAdapter(this.Company, Convert.ToInt32(newDocEntry));
        }

        /// <summary>
        /// Adds the multiple activity.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="objectType">Type of the object.</param>
        /// <param name="userSign">The user sign.</param>
        /// <returns>True if the multiple activity is added successfully</returns>
        public bool AddMultipleActivity(int document, string objectType, int userSign)
        {
            return this.AddMultipleActivity(Convert.ToInt32(this.Document.DocEntry), document, objectType, userSign);
        }

        /// <summary>
        /// Adds the multiple activity.
        /// </summary>
        /// <param name="activityCode">The activity code.</param>
        /// <param name="document">The document.</param>
        /// <param name="objectType">Type of the object.</param>
        /// <param name="userSign">The user sign.</param>
        /// <returns>True if the multiple activity is added successfully</returns>
        public bool AddMultipleActivity(int activityCode, int document, string objectType, int userSign)
        {
            this.Task = "Adding a multiple activity";

            // Check if the activity already exists
            var stringBuilder = new StringBuilder();
            stringBuilder.AppendLine(@"SELECT 1 FROM [@XX_DOCACTIVITIES] ");
            stringBuilder.AppendFormat(@"WHERE [U_ObjType] = '{0}' ", objectType);
            stringBuilder.AppendFormat(@"AND [U_DocEntry] = '{0}' ", document);
            stringBuilder.AppendFormat(@"AND [U_ActEntry] = '{0}' ", activityCode);

            using (var rs = new RecordsetAdapter(this.Company, stringBuilder.ToString()))
            {
                if (rs.EoF)
                {
                    // Add multiple activity
                    int docEntry;
                    string code;
                    this.Exchanger.GetNextCode("DocumentActivities", string.Empty, 5, out code, out docEntry);

                    var createdDate = DateTime.Today.ToString("yyyyMMdd");
                    var createdTime = DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00");

                    stringBuilder = new StringBuilder();
                    stringBuilder.AppendLine(@"INSERT INTO [@XX_DOCACTIVITIES]([Code],[Name],[DocEntry],");
                    stringBuilder.AppendLine(@"[Canceled],[Object],[LogInst],[UserSign],");
                    stringBuilder.AppendLine(@"[Transfered],[CreateDate],[CreateTime],");
                    stringBuilder.AppendLine(@"[UpdateDate],[UpdateTime],[DataSource],");
                    stringBuilder.AppendLine(@"[U_ObjType],[U_DocEntry],[U_ActEntry])");
                    stringBuilder.AppendFormat(@"VALUES('{0}','',{1},'N','DocumentActivities',Null,{2},{3}", code, docEntry, userSign, Environment.NewLine);
                    stringBuilder.AppendFormat(@"'N','{0}',{4}, Null,Null,'I','{1}','{2}','{3}')", createdDate, objectType, document, activityCode, createdTime);
                    RecordsetAdapter.Execute(this.Company, stringBuilder.ToString());
                }
            }

            return true;
        }

        /// <summary>
        /// Deletes the multiple activity.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="objectType">Type of the object.</param>
        /// <returns>True if the mulpiple activity is successfully deleted.</returns>
        public bool DeleteMultipleActivity(int document, string objectType)
        {
            this.Task = string.Format("Deleting multiple activity from {0} of type {1}", document, objectType) ;

            RecordsetAdapter.Execute(this.Company, string.Format("DELETE FROM [@XX_DOCACTIVITIES] WHERE [U_ObjType] = '{0}' AND U_ActEntry = '{1}' AND U_DocEntry = '{2}'", objectType, document, this.Document.ContactCode));

            return true;
        }

        /// <summary>
        /// Transfers all Userdefined fields to adapter.
        /// </summary>
        /// <param name="targetAdapter">The target adapter.</param>
        public void TransferFieldsToAdapter(ActivityAdapter targetAdapter)
        {
            var documentUserFields = this.Document.UserFields;
            try
            {
                var documentUserFieldsFields = documentUserFields.Fields;

                try
                {
                    // Transfer all the custom fields value
                    for (int fieldIndex = 0; fieldIndex < documentUserFieldsFields.Count; fieldIndex++)
                    {
                        targetAdapter.SetUserDefinedField(fieldIndex, this.GetUserDefinedField(fieldIndex));
                    }
                }
                finally
                {
                    COMHelper.Release(ref documentUserFieldsFields);
                    documentUserFieldsFields = null;
                }
            }
            finally
            {
                COMHelper.Release(ref documentUserFields);
                documentUserFields = null;
            }
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
        /// Gets the user defined field from the Document.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <returns>The value of the field</returns>
        public object GetUserDefinedField(object fieldIndex)
        {
            return COMHelper.UserDefinedFieldValue(this.Document.UserFields, fieldIndex);
        }

        /// <summary>
        /// Releases the Documents COM Object.
        /// </summary>
        protected override void Release()
        {
            if (this.Document != null)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(this.Document);
                }
                finally
                {
                    this.Document = null;
                    GC.Collect();
                }
            }
        }
        #endregion Methods
    }
}
