//-----------------------------------------------------------------------
// <copyright file="MessageAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters
{
    #region Using Directive(s)

    using SAPbobsCOM;
    using Helpers;
    using Model.Exceptions;

    #endregion Using Directive(s)

    public class MessageAdapter : SapObjectAdapter
    {
        #region Private Member(s)

        private Recipients _recipients;

        private Attachments _attachments;

        #endregion Private Member(s)

        #region Properties

        /// <summary>
        /// Gets or sets the message.
        /// </summary>
        /// <value>The message.</value>
        public Messages Message
        {
            get; 
            private set;
        }

        /// <summary>
        /// Gets or sets the subject.
        /// </summary>
        /// <value>The subject.</value>
        public string Subject
        {
            get
            {
                return this.Message.Subject;
            }
            set { this.Message.Subject = value; }
        }

        /// <summary>
        /// Gets or sets the body.
        /// </summary>
        /// <value>The body.</value>
        public string Body
        {
            get { return this.Message.MessageText; }
            set { this.Message.MessageText = value; }
        }

        /// <summary>
        /// Gets or sets the priority.
        /// </summary>
        /// <value>The priority.</value>
        public BoMsgPriorities Priority
        {
            get { return this.Message.Priority; }
            set { this.Message.Priority = value; }
        }

        /// <summary>
        /// Gets the message recipients.
        /// </summary>
        /// <value>The message recipients.</value>
        protected Recipients MessageRecipients
        {
            get { return this._recipients ?? (this._recipients = this.Message.Recipients); }
        }

        /// <summary>
        /// Gets the attachments.
        /// </summary>
        /// <value>The attachments.</value>
        protected Attachments Attachments
        {
            get { return this._attachments ?? (this._attachments = this.Message.Attachments); }   
        }

        #endregion Properties

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="SapObjectAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        public MessageAdapter(Company company) : base(company)
        {
            this.Message = (Messages)this.Company.GetBusinessObject(BoObjectTypes.oMessages);

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SapObjectAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="message">The message</param>
        public MessageAdapter(Company company, Messages message)
            : base(company)
        {
            this.Message = message;
        }

        #endregion Constructor(s)

        #region Method(s)

        /// <summary>
        /// Adds an internal recipient.
        /// </summary>
        /// <param name="userCode">The user code.</param>
        public void AddInternalRecipient(string userCode)
        {
            if (this.MessageRecipients.Count > 0)
            {
                this.MessageRecipients.Add();
            }

            this.MessageRecipients.UserCode = userCode;
            this.MessageRecipients.SendInternal = BoYesNoEnum.tYES;
            this.MessageRecipients.SendEmail = BoYesNoEnum.tNO;
            this.MessageRecipients.SendFax = BoYesNoEnum.tNO;
            this.MessageRecipients.SendSMS = BoYesNoEnum.tNO;
        }


        /// <summary>
        /// Adds an email recipient.
        /// </summary>
        /// <param name="emailAddress">The email address.</param>
        public void AddEmailRecipient(string emailAddress)
        {
            if (this.MessageRecipients.Count > 0)
            {
                this.MessageRecipients.Add();
            }

            this.MessageRecipients.EmailAddress = emailAddress;
            this.MessageRecipients.SendInternal = BoYesNoEnum.tNO;
            this.MessageRecipients.SendEmail = BoYesNoEnum.tYES;
            this.MessageRecipients.SendFax = BoYesNoEnum.tNO;
            this.MessageRecipients.SendSMS = BoYesNoEnum.tNO;
        }

        /// <summary>
        /// Adds an internal & email recipient.
        /// </summary>
        /// <param name="userCode">The user code.</param>
        /// <param name="emailAddress">The email address.</param>
        public void AddRecipient(string userCode, string emailAddress)
        {
            if (this.MessageRecipients.Count > 0)
            {
                this.MessageRecipients.Add();
            }

            this.MessageRecipients.UserCode = userCode;
            this.MessageRecipients.EmailAddress = emailAddress;
            this.MessageRecipients.SendInternal = BoYesNoEnum.tYES;
            this.MessageRecipients.SendEmail = BoYesNoEnum.tYES;
            this.MessageRecipients.SendFax = BoYesNoEnum.tNO;
            this.MessageRecipients.SendSMS = BoYesNoEnum.tNO;
            
        }

        /// <summary>
        /// Adds an attachment to the message.
        /// </summary>
        /// <param name="attachment">The attachment.</param>
        public void AddAttachment(string attachment)
        {
            if (this.Attachments.Count > 0)
            {
                this.Attachments.Add();
            }

            this.Attachments.Item(0).FileName = attachment;
        }

        /// <summary>
        /// Sends the message.
        /// </summary>
        public void SendMessage()
        {
            if (this.Message.Add() != 0)
            {
                throw new SapException(this.Company);
            }
        }

        #region SapObjectAdapter implementation

        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected override void Release()
        {
            if (this._attachments != null)
            {
                COMHelper.Release(ref _attachments);
            }

            if (this._recipients != null)
            {
                COMHelper.Release(ref _recipients);
            }

            if (this.Message != null)
            {
                COMHelper.ReleaseOnly(this.Message);
                this.Message = null;
            }
        }

        #endregion SapObjectAdapter implementation

        #endregion Method(s)
    }
}
