//-----------------------------------------------------------------------
// <copyright file="AddOn.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.UI
{
    #region Usings
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    
    using System.Threading;
    using System.Windows.Forms;
    using DI.Components;
    using DI.Helpers;
    using Helpers;
    using Model.EventArguments;
    using SAPbouiCOM;
    using Windows;
    using Application = SAPbouiCOM.Application;
    using Company = SAPbobsCOM.Company;
    using Form = System.Windows.Forms.Form;
    using IForm = Model.Interfaces.IForm;
    using StatusBar = Windows.StatusBar;
    using System.Xml;
    using B1C.SAP.DI.BusinessAdapters.Administration;
    using B1C.SAP.DI.Model.Exceptions;
    using System.Text.RegularExpressions; 
    #endregion

    #region Delegates

    /// <summary>
    /// Handles Menu Event
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    /// <param name="bubbleEvent">The bubble event</param>
    public delegate void MenuEventsHandler(object sender, ref MenuEvent e, out bool bubbleEvent);

    /// <summary>
    /// Handles Menu Click
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void MenuClickHandler(object sender, MenuEventArgs e);

    /// <summary>
    /// Handles Form Event
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void NewFormCreatedHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles The database update event
    /// </summary>
    /// <param name="sender">The event sender.</param>
    public delegate void UpdateDatabaseHandler(object sender);
    #endregion

    /// <summary>
    /// Class name AddOn
    /// Defined in namespace B1C.SAP.UI
    /// </summary>
    public abstract class AddOn
    {
        #region Fields
        /// <summary>
        /// Connection String
        /// </summary>
        private readonly string _connectionString = string.Empty;

        /// <summary>
        /// Current Application Object
        /// </summary>
        private Application _application;

        /// <summary>
        /// Current Company
        /// </summary>
        private Company _company;

        /// <summary>
        /// True if we are connected to SAP
        /// </summary>
        private bool _connected;

        /// <summary>
        /// The Form Factory
        /// </summary>
        private FormFactory _formFactory;

        /// <summary>
        /// The forms collection
        /// </summary>
        private Dictionary<string, IForm> _forms;

        /// <summary>
        /// True if you want to log the information messages controlled
        /// by the application configuration file
        /// </summary>
        private bool _logInfo;

        /// <summary>
        /// This is the default add on name
        /// </summary>
        private string _name = "B1C AddOn";

        /// <summary>
        /// Status bar
        /// </summary>
        private StatusBar _statusBar;

        /// <summary>
        /// SQL Connection
        /// </summary>
        private System.Data.SqlClient.SqlConnection _connection;

        /// <summary>
        /// Tool bar
        /// </summary>
        private Constants.Toolbar _toolbar;

        /// <summary>
        /// The Application Menu
        /// </summary>
        private Constants.Menu _menu;
        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="AddOn"/> class.
        /// </summary>
        protected AddOn()
        {
            // Make sure that only one instance of the add-on is running

            string addOnProcessName = Process.GetCurrentProcess().MainModule.ModuleName;

            int currentProcessId = Process.GetCurrentProcess().Id;

            // Remove the Process Name without extension
            addOnProcessName = Path.GetFileNameWithoutExtension(addOnProcessName);

            // Makes a list of processes with the same name
            Process[] process = Process.GetProcessesByName(addOnProcessName);

            // Kill any other process that is not the current
            for (int processIndex = process.Length - 1; processIndex == 0; processIndex--)
            {
                if (process[processIndex].Id != currentProcessId)
                {
                    process[processIndex].Kill();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }


            // Get the log info
            try
            {
                ConfigurationHelper.Load("LogInformation", out _logInfo);
            }
            catch
            {
            }

#if DEBUG
            _connectionString =
                "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
#else
    // Get the log info
            if (Environment.GetCommandLineArgs().Length > 1 )
            {
                _connectionString = Environment.GetCommandLineArgs().GetValue(1).ToString();
            }
#endif

            if (!Connect())
            {
                System.Windows.Forms.MessageBox.Show("Unabled to connect.\rNo SAP Business One application was found.",
                                                     "Connecting to SAP Business One", MessageBoxButtons.OK,
                                                     MessageBoxIcon.Error);
                Terminate();
            }
            else
            {
                // events handled by SBO_Application_AppEvent
                Application.AppEvent += ApplicationAppEvent;
                Application.StatusBarEvent += StatusBarEvent;
                Application.ItemEvent += ApplicationItemEvent;
                Application.MenuEvent += ApplicationMenuEvent;
                Application.RightClickEvent += ApplicationRightClickEvent;
            }

            this.SetLogPath();
            LogInformation("Base Add-on Constructor: End Constructor");
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AddOn"/> class.
        /// </summary>
        /// <param name="addonNamespace">The addon namespace.</param>
        /// <param name="addonId">The addon id.</param>
        protected AddOn(string addonNamespace, string addonId)
        {
            // Get the log info
            try
            {
                ConfigurationHelper.Load("LogInformation", out _logInfo);
            }
            catch
            {
            }

#if DEBUG
            _connectionString =
                "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
#else
    // Get the log info
            if (Environment.GetCommandLineArgs().Length > 1)
            {
                _connectionString = Environment.GetCommandLineArgs().GetValue(1).ToString();
            }
#endif

            if (!Connect())
            {
                System.Windows.Forms.MessageBox.Show(
                    "Unabled to connect.\rSAP Business One must be running for this addon to start.",
                    "SAP Business One Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Terminate();
            }
            else
            {

                AddonId = addonId;
                Namespace = addonNamespace;

                // this.ApplyPatches();

                // events handled by SBO_Application_AppEvent
                Application.AppEvent += ApplicationAppEvent;
                Application.StatusBarEvent += StatusBarEvent;
                Application.ItemEvent += ApplicationItemEvent;
                Application.MenuEvent += ApplicationMenuEvent;
                Application.RightClickEvent += ApplicationRightClickEvent;

                this.RegisterEvents();
                this.UpdateSchema();
            }

            this.SetLogPath();
            LogInformation("Base Add-on Constructor: End Constructor");
        }

        #endregion Constructors

        #region Events

        /// <summary>
        /// Occurs when [item click handler].
        /// </summary>
        public event MenuClickHandler MenuClick;

        /// <summary>
        /// Occurs when [menu event].
        /// </summary>
        [Obsolete("This method should be avoided. Only for back compatibility with the form controller implementation.")]
        public event MenuEventsHandler MenuEvents;

        /// <summary>
        /// Occurs when [new form created].
        /// </summary>
        public event NewFormCreatedHandler NewFormCreated;

        /// <summary>
        /// Occurs when [update database].
        /// </summary>
        public event UpdateDatabaseHandler UpdateDatabase;
        #endregion

        #region Properties

        /// <summary>
        /// Gets the tool bar.
        /// </summary>
        /// <value>The tool bar.</value>
        public Constants.Toolbar ToolBar
        {
            get
            {
                if (this._toolbar == null)
                {
                    this._toolbar = new Constants.Toolbar();
                }

                return _toolbar;
            }
        }


        /// <summary>
        /// Gets the menu.
        /// </summary>
        /// <value>The application menu.</value>
        public Constants.Menu Menu
        {
            get
            {
                if (this._menu == null)
                {
                    this._menu = new Constants.Menu();
                }

                return _menu;
            }
        }

        /// <summary>
        /// Gets or sets the name space.
        /// </summary>
        /// <value>The name space.</value>
        public string Namespace { get; set; }

        /// <summary>
        /// Gets or sets the addon id.
        /// </summary>
        /// <value>The addon id.</value>
        public string AddonId { get; set; }

        /// <summary>
        /// Gets the status bar.
        /// </summary>
        /// <value>The status bar.</value>
        public StatusBar StatusBar
        {
            get { return _statusBar ?? (_statusBar = new StatusBar(this)); }
        }

        /// <summary>
        /// Gets the SAP application instance.
        /// </summary>
        /// <value>The SAP application.</value>
        public Application Application
        {
            get { return _application; }
        }

        /// <summary>
        /// Gets the company instance.
        /// </summary>
        /// <value>The company.</value>
        public Company Company
        {
            get { return _company; }
        }

        /// <summary>
        /// Gets a value indicating whether this <see cref="AddOn"/> is connected.
        /// </summary>
        /// <value><c>true</c> if connected; otherwise, <c>false</c>.</value>
        public bool Connected
        {
            get { return _connected; }
        }

        /// <summary>
        /// Gets the connection string.
        /// </summary>
        /// <value>The connection string.</value>
        public string ConnectionString
        {
            get { return _connectionString; }
        }

        /// <summary>
        /// Gets the location.
        /// </summary>
        /// <value>The location.</value>
        public string Location
        {
            get
            {
                string returnValue = string.Empty;
                string errorMessage = string.Empty;
                try
                {
#if DEBUG
                    returnValue = System.Windows.Forms.Application.StartupPath;
                    returnValue = returnValue.Replace("bin\\Debug", "Resources");
#else
                    returnValue = System.Windows.Forms.Application.StartupPath;
#endif
                }
                catch (Exception ex)
                {
                    errorMessage = ex.Message;
                }

                if (!string.IsNullOrEmpty(errorMessage))
                {
                    System.Windows.Forms.MessageBox.Show(errorMessage, "Error Occurred", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                if (!returnValue.EndsWith("\\"))
                {
                    returnValue += "\\";
                }

                var dir = new DirectoryInfo(returnValue);
                if (!dir.Exists)
                {
                    dir.Create();
                }

                return returnValue;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to log information messages.
        /// </summary>
        /// <value><c>true</c> if [log info]; otherwise, <c>false</c>.</value>
        public bool LogInfo
        {
            get { return _logInfo; }

            set { _logInfo = value; }
        }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>The name of the add on.</value>
        public string Name
        {
            get { return _name; }

            set { _name = value; }
        }

        /// <summary>
        /// Gets the folder.
        /// </summary>
        /// <value>The folder class.</value>
        public SystemFolder Folder
        {
            get { return new SystemFolder(_company); }
        }

        /// <summary>
        /// Gets the report user.
        /// </summary>
        /// <value>The report user.</value>
        public ReportUser ReportUser
        {
            get { return new ReportUser(); }
        }

        /// <summary>
        /// Gets or sets the action cancelled message.
        /// </summary>
        /// <value>The action cancelled message.</value>
        public string ActionCancelledMessage { get; set; }

        /// <summary>
        /// Gets the form factory.
        /// </summary>
        /// <value>The form factory.</value>
        public FormFactory FormFactory
        {
            get { return _formFactory ?? (_formFactory = new FormFactory(this)); }
        }

        /// <summary>
        /// Gets the forms.
        /// </summary>
        /// <value>The forms.</value>
        public Dictionary<string, IForm> Forms
        {
            get { return _forms ?? (_forms = new Dictionary<string, IForm>()); }
        }

        /// <summary>
        /// Gets or sets the form count.
        /// </summary>
        /// <value>The form count.</value>
        protected internal int FormCount { get; set; }

        public System.Data.SqlClient.SqlConnection Connection
        {
            get 
            {
                try
                {
                    if (this._connection == null)
                    {
                        string connectionString = string.Format("Data Source={0}; Initial Catalog={1}; Trusted_Connection=True;", this.Company.Server, this.Company.CompanyDB);
                        this._connection = new System.Data.SqlClient.SqlConnection(connectionString);
                    }

                    if ((this._connection.State == System.Data.ConnectionState.Closed) || (this._connection.State == System.Data.ConnectionState.Broken))
                    {
                        this._connection.Open();
                    }
                }
                catch (Exception ex)
                {
                    // Try to close the connection
                    if (this._connection != null)
                    {
                        this._connection.Dispose();
                    }

                    throw ex;
                }

                return this._connection; 
            }

            set 
            { 
                this._connection = value; 
            }
        }

        /// <summary>
        /// Gets the active form.
        /// </summary>
        /// <value>The active form.</value>
        public IForm ActiveForm
        {
            get { return this.Forms[this.Application.Forms.ActiveForm.UniqueID]; }
        }
        #endregion Properties

        #region Methods

        /// <summary>
        /// Updates the schema.
        /// </summary>
        private void UpdateSchema()
        {
            bool databaseUpdated;
            ConfigurationHelper.Load("DatabaseUpdated", out databaseUpdated);
            if (!databaseUpdated)
            {
                // Raise the Update database event
                this.OnUpdateDatabase();
                ConfigurationHelper.Save("DatabaseUpdated", true);
            }
        }

        /// <summary>
        /// Sends the email.
        /// </summary>
        /// <param name="emailSubject">The email subject.</param>
        /// <param name="emailBody">The email body.</param>
        /// <param name="internalRecipients">The internal recipients.</param>
        /// <param name="externalRecipients">The external recipients.</param>
        /// <param name="priority">The priority.</param>
        public void SendEmail(string emailSubject, string emailBody, string internalRecipients, string externalRecipients, SAPbobsCOM.BoMsgPriorities priority)
        {
            var message = (SAPbobsCOM.Messages)this.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);

            message.Subject = emailSubject;
            message.MessageText = emailBody;

            // Add each recipient in the comma separated list
            var internalRecipientList = internalRecipients.Split(',');
            var recipientCount = 0;

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
                    message.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                    recipientCount++;
                }
            }

            var externalRecipientList = externalRecipients.Split(',');

            foreach (var recipient in externalRecipientList)
            {
                var trimmedRecipient = recipient.Trim();
                if (string.IsNullOrEmpty(trimmedRecipient)) continue;

                if (recipientCount > 0)
                {
                    message.Recipients.Add();
                    message.Recipients.SetCurrentLine(recipientCount);
                }

                message.Recipients.UserCode = UserAdapter.Current(this.Company).UserCode;
                message.Recipients.EmailAddress = trimmedRecipient;
                message.Recipients.SendEmail = SAPbobsCOM.BoYesNoEnum.tYES;
                recipientCount++;
            }

            message.Priority = priority;

            if (message.Add() != 0)
            {
                throw new SapException(this.Company);
            }
        }

        /// <summary>
        /// Adds the form.
        /// </summary>
        /// <param name="form">The form being added.</param>
        public void AddForm(IForm form)
        {
            if (Forms.ContainsKey(form.FormId))
            {
                Debug.Print(string.Format("Form {0} of form type {1} already in collection.", form.FormId, form.FormTypeId));

                form.Freeze(false);
                form.Dispose();
                form = null;
                return;
            }

            Debug.Print(string.Format("Form {0} of form type {1} added in collection.", form.FormId, form.FormTypeId));
            // form.Initialize();
            CaptureEvents(form.FormTypeId);
            form.ApplyFilters();
            form.Disposed += DisposedForm;
            Forms.Add(form.FormId, form);
            form.Refresh();

            form.Freeze(false);
        }

        /// <summary>
        /// Messages the box.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="buttons">
        /// The buttons.
        /// </param>
        /// <returns>
        /// The result from the dialog button pressed
        /// </returns>
        public DialogResult MessageBox(string message, MessageBoxButtons buttons)
        {
            int result;
            if (buttons == MessageBoxButtons.AbortRetryIgnore)
            {
                result = Application.MessageBox(message, 1, "Abort", "Retry", "Ignore");
                if (result == 1)
                {
                    return DialogResult.Abort;
                }

                if (result == 2)
                {
                    return DialogResult.Retry;
                }

                if (result == 3)
                {
                    return DialogResult.Ignore;
                }
            }
            else if (buttons == MessageBoxButtons.OK)
            {
                result = Application.MessageBox(message, 1, "Ok", string.Empty, string.Empty);
                if (result == 1)
                {
                    return DialogResult.OK;
                }
            }
            else if (buttons == MessageBoxButtons.OKCancel)
            {
                result = Application.MessageBox(message, 1, "Ok", "Cancel", string.Empty);
                if (result == 1)
                {
                    return DialogResult.OK;
                }

                if (result == 2)
                {
                    return DialogResult.Cancel;
                }
            }
            else if (buttons == MessageBoxButtons.RetryCancel)
            {
                result = Application.MessageBox(message, 1, "Retry", "Cancel", string.Empty);
                if (result == 1)
                {
                    return DialogResult.Retry;
                }

                if (result == 2)
                {
                    return DialogResult.Cancel;
                }
            }
            else if (buttons == MessageBoxButtons.YesNo)
            {
                result = Application.MessageBox(message, 1, "Yes", "No", string.Empty);
                if (result == 1)
                {
                    return DialogResult.Yes;
                }

                if (result == 2)
                {
                    return DialogResult.No;
                }
            }
            else if (buttons == MessageBoxButtons.YesNoCancel)
            {
                result = Application.MessageBox(message, 1, "Yes", "No", "Cancel");
                if (result == 1)
                {
                    return DialogResult.Yes;
                }

                if (result == 2)
                {
                    return DialogResult.No;
                }

                if (result == 3)
                {
                    return DialogResult.Cancel;
                }
            }

            return DialogResult.OK;
        }

        /// <summary>
        /// Messages the box.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <returns>
        /// The result from the dialog button pressed
        /// </returns>
        public DialogResult MessageBox(string message)
        {
            return MessageBox(message, MessageBoxButtons.OK);
        }

        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="ex">The exception.</param>
        public void LogError(Exception ex)
        {
            LogHandler.LogError(ex);

            if (ex.Message.Contains("HRESULT"))
            {
                Terminate();
            }
        }

        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="ex">The exception.</param>
        /// <param name="message">The message.</param>
        public void LogError(Exception ex, string message)
        {
            LogHandler.LogError(ex, message);

            if (ex.Message.Contains("HRESULT"))
            {
                Terminate();
            }
        }

        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="ex">The exception.</param>
        /// <param name="messages">The messages.</param>
        public void LogError(Exception ex, ArrayList messages)
        {
            LogHandler.LogError(ex, messages);

            if (ex.Message.Contains("HRESULT"))
            {
                Terminate();
            }
        }

        /// <summary>
        /// Logs the information.
        /// </summary>
        /// <param name="message">The message.</param>
        public void LogInformation(string message)
        {
            if (_logInfo)
            {
                LogHandler.LogInformation(message);
            }
        }

        /// <summary>
        /// Logs the information.
        /// </summary>
        /// <param name="messages">The messages.</param>
        public void LogInformation(ArrayList messages)
        {
            if (_logInfo)
            {
                LogHandler.LogInformation(messages);
            }
        }

        /// <summary>
        /// Terminates this instance.
        /// </summary>
        public void Terminate()
        {
            LogInformation("Base Add-on Terminate: Begin termination.");
            if (_company != null)
            {
                _company.Disconnect();
            }

            if (this._connection != null)
            {
                if (this._connection.State == System.Data.ConnectionState.Open)
                {
                    this._connection.Close();
                }
                this._connection.Dispose();
            }

            LogInformation("Base Add-on Terminate: Company Disconnected.");
            _application = null;
            LogInformation("Base Add-on Terminate: Application cleared.");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            LogInformation("Base Add-on Terminate: Garbage Collector cleared.");

            // terminating the Add On
            System.Windows.Forms.Application.Exit();
            Environment.Exit(0);
        }

        /// <summary>
        /// Shows the specified windows form.
        /// </summary>
        /// <param name="windowsForm">The windows form.</param>
        public void Show(Form windowsForm)
        {
            System.Windows.Forms.Application.Run(windowsForm);
        }

        /// <summary>
        /// Singles the threaded.
        /// </summary>
        /// <param name="method">The method to run.</param>
        public void SingleThreaded(ThreadStart method)
        {
            var thread = new Thread(method);
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
#if RELEASE
            thread.Join();
#endif
        }

        /// <summary>
        /// Captures the events.
        /// </summary>
        /// <param name="formType">Type of the form.</param>
        public void CaptureEvents(string formType)
        {
            AddFilter(BoEventTypes.et_FORM_UNLOAD, formType);
            AddFilter(BoEventTypes.et_FORM_LOAD, formType);
        }

        /// <summary>
        /// Releases the events.
        /// </summary>
        public abstract void UnregisterEvents();

        /// <summary>
        /// Handles the events.
        /// </summary>
        public abstract void RegisterEvents();

        /// <summary>
        /// Adds the filter if not already trapped by the add-on.
        /// </summary>
        /// <param name="eventType">Type of the event.</param>
        /// <param name="formType">Type of the form.</param>
        internal void AddFilter(BoEventTypes eventType, string formType)
        {
            bool filtersChanged = false;
            EventFilters filters = this.Application.GetFilter();

            EventFilter filter = null;
            foreach (EventFilter addFilter in filters)
            {
                //if (addFilter.EventType == BoEventTypes.et_ALL_EVENTS)
                //{
                //    allIsPresent = true;

                //}

                if (addFilter.EventType == eventType)
                {
                    filter = addFilter;
                    break;
                }
            }

            if (filter == null)
            {
                filters.Add(eventType);
                filter = filters.Item(filters.Count - 1);
                filtersChanged = true;
            }

            if (!string.IsNullOrEmpty(formType))
            {
                if (formType.StartsWith(string.Format("{0}_{1}", this.Namespace, this.AddonId)))
                {
                    // This is a custom Form
                    bool foundFilter = false;
                    foreach (FormType addForm in filter)
                    {
                        if (addForm.StringValue == formType)
                        {
                            foundFilter = true;
                            break;
                        }
                    }

                    if (!foundFilter)
                    {
                        Debug.WriteLine(string.Format("Added filter {0} on form {1}", eventType, formType));
                        filter.AddEx(formType);
                        filtersChanged = true;
                    }
                }
                else
                {
                    bool foundFilter = false;
                    foreach (FormType addForm in filter)
                    {
                        if (addForm.Value == Convert.ToInt32(formType))
                        {
                            foundFilter = true;
                            break;
                        }
                    }

                    if (!foundFilter)
                    {
                        Debug.WriteLine(string.Format("Added filter {0} on form {1}", eventType, formType));
                        filter.Add(Convert.ToInt32(formType));
                        filtersChanged = true;
                    }
                }
            }

            //if (allIsPresent)
            //{
            //    filtersChanged = true;
            //    filters.Remove(SAPbouiCOM.BoEventTypes.et_ALL_EVENTS);

            //    // Add the Menu click
            //    try
            //    {
            //        filters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            //    }
            //    catch
            //    {
            //    }
            //}

            if (filtersChanged)
            {
                this.Application.SetFilter(filters);
            }
        }
        /// <summary>
        /// Connects this instance.
        /// </summary>
        /// <returns>True if it is connected</returns>
        protected bool Connect()
        {
            try
            {
                LogInformation("Base Add-on Connect: connection string ==>" + _connectionString);

                // Validate connection string
                if (string.IsNullOrEmpty(_connectionString))
                {
                    return false;
                }

                var userInterfaceApi = new SboGuiApi();
                LogInformation("Base Add-on Connect: SboGuiAPi loaded.");

                // Connect to the UI
                userInterfaceApi.Connect(_connectionString);
                LogInformation("Base Add-on Connect: SboGuiAPi connected.");

                // Initialize the application
                _application = userInterfaceApi.GetApplication(-1);
                LogInformation("Base Add-on Connect: Application retrieved.");

                try
                {
                    _company = (Company) Application.Company.GetDICompany();
                    LogInformation("Base Add-on Connect: Company retrieved.");
                }
                catch (Exception ex)
                {
                    this.HandleException(ex);
                    this.Terminate();
                }

                if (Application != null)
                {
                    /***************************************
                    * Raise the Initialization Events
                    ***************************************/
                    this.LoadMenus();
                }

                _connected = Application != null;
                LogInformation("Base Add-on Connect: Connected successfully.");
            }
            catch
            {
                _connected = false;
                LogInformation("Base Add-on Connect: Connected failure.");
            }

            return _connected;
        }

        /// <summary>
        /// Reset the connection.
        /// </summary>
        /// <returns>True if re-connects successfully</returns>
        protected bool ReConnect()
        {
            if (_company != null)
            {
                _company.Disconnect();
            }

            _application = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return Connect();
        }

        /// <summary>
        /// Raises the ItemClick event.
        /// </summary>
        /// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnMenuClick(MenuEventArgs e)
        {
            try
            {
                if (MenuClick != null)
                {
                    MenuClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.HandleException(ex);
            }
        }

        /// <summary>
        /// Called when [menu events].
        /// </summary>
        /// <param name="e">The e.</param>
        /// <param name="bubbleEvent">if set to <c>true</c> [bubble event].</param>
        protected virtual void OnMenuEvents(ref MenuEvent e, out bool bubbleEvent)
        {
            bubbleEvent = true;

            try
            {
                if (MenuEvents != null)
                {
                    MenuEvents(this, ref e, out bubbleEvent);
                }
            }
            catch (Exception ex)
            {
                this.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the AfterFormLoad event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnNewFormCreated(FormEventArgs e)
        {
            try
            {
                if (NewFormCreated != null)
                {
                    NewFormCreated(this, e);
                }
            }
            catch (Exception ex)
            {
                this.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the UpdateDatabase event.
        /// </summary>
        protected virtual void OnUpdateDatabase()
        {
            try
            {
                if (UpdateDatabase != null)
                {
                    UpdateDatabase(this);
                }
            }
            catch (Exception ex)
            {
                this.HandleException(ex);
            }
        }

        /// <summary>
        /// Application events.
        /// </summary>
        /// <param name="eventType">Type of the event.</param>
        private void ApplicationAppEvent(BoAppEventTypes eventType)
        {
            switch (eventType)
            {
                case BoAppEventTypes.aet_ShutDown:
                    // terminating the Add On
                    Terminate();
                    break;
                case BoAppEventTypes.aet_ServerTerminition:
                    // terminating the Add On
                    Terminate();
                    break;
            }
        }

        /// <summary>
        /// Statuses the bar event.
        /// </summary>
        /// <param name="displayMessage">The display message.</param>
        /// <param name="messageType">Type of the message.</param>
        private void StatusBarEvent(string displayMessage, BoStatusBarMessageType messageType)
        {
            switch (messageType)
            {
                case BoStatusBarMessageType.smt_Error:
                    // terminating the Add On
                    if (displayMessage.Contains("UI_API -7780"))
                    {
                        Application.StatusBar.SetText(string.Empty, BoMessageTime.bmt_Short,
                                                      BoStatusBarMessageType.smt_None);
                    }
                    else if (displayMessage.Contains("dbo.OGFL"))
                    {
                        Application.StatusBar.SetText(string.Empty, BoMessageTime.bmt_Short,
                                                      BoStatusBarMessageType.smt_None);
                    }

                    break;
            }
        }

        /// <summary>
        /// Applications the right click event.
        /// </summary>
        /// <param name="contextMenuInfo">The context menu info.</param>
        /// <param name="bubbleEvent">if set to <c>true</c> [bubble event].</param>
        private void ApplicationRightClickEvent(ref ContextMenuInfo contextMenuInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            this.RunRightClick(ref contextMenuInfo, out bubbleEvent);
        }

        /// <summary>
        /// Applications the item event.
        /// </summary>
        /// <param name="formUID">The form UID.</param>
        /// <param name="itemEvent">The item event.</param>
        /// <param name="bubbleEvent">if set to <c>true</c> [bubble event].</param>
        private void ApplicationItemEvent(string formUID, ref ItemEvent itemEvent, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (!Forms.ContainsKey(formUID))
            {
                // Raised the NewFormCreated event
                OnNewFormCreated(new FormEventArgs(itemEvent.FormTypeEx, formUID, string.Empty));
            }

            if (!string.IsNullOrEmpty(ActionCancelledMessage))
            {
                Application.StatusBar.SetText(ActionCancelledMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                ActionCancelledMessage = string.Empty;
            }
        }

        /// <summary>
        /// Applications the menu event.
        /// </summary>
        /// <param name="menuEvent">The menu event.</param>
        /// <param name="bubbleEvent">if set to <c>true</c> [bubble event].</param>
        private void ApplicationMenuEvent(ref MenuEvent menuEvent, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (!menuEvent.BeforeAction)
            {
                AddFilter(BoEventTypes.et_MENU_CLICK, string.Empty);
                OnMenuClick(new MenuEventArgs(menuEvent.MenuUID));
            }

            bool cancelEvent;
            string cancelMessage;
            RunToolBarClick(menuEvent.MenuUID, menuEvent.BeforeAction, out cancelEvent, out cancelMessage);

            if (cancelEvent)
            {
                bubbleEvent = false;
                if (!string.IsNullOrEmpty(cancelMessage))
                {
                    ActionCancelledMessage = cancelMessage;
                    this.StatusBar.ErrorMessage = ActionCancelledMessage;
                }
            }
            else
            {
                OnMenuEvents(ref menuEvent, out bubbleEvent);
            }
        }

        /// <summary>
        /// Disposeds the form.
        /// </summary>
        /// <param name="sender">The sender.</param>
        private void DisposedForm(object sender)
        {
            string formId = ((Windows.Form) sender).FormId;
            if (!Forms.ContainsKey(formId)) return;

            Forms[formId].Disposed -= DisposedForm;
            Forms[formId] = null;
            Forms.Remove(formId);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Runs the tool bar click.
        /// </summary>
        /// <param name="menuId">The menu id.</param>
        /// <param name="beforeEvent">if set to <c>true</c> [before event].</param>
        /// <param name="cancelEvent">if set to <c>true</c> [cancel event].</param>
        /// <param name="cancelMessage">The cancel message.</param>
        private void RunToolBarClick(string menuId, bool beforeEvent, out bool cancelEvent, out string cancelMessage)
        {
            cancelEvent = false;
            cancelMessage = string.Empty;

            if (!Forms.ContainsKey(Application.Forms.ActiveForm.UniqueID))
            {
                return;
            }

            IForm form = Forms[Application.Forms.ActiveForm.UniqueID];
            if (this.ToolBar.ContainsItem(menuId))
            {
                if (Forms.ContainsKey(Application.Forms.ActiveForm.UniqueID))
                {
                    form.ToolBarClicked(menuId, beforeEvent, out cancelEvent, out cancelMessage);
                }
            }

            if (!cancelEvent)
            {
                form.MenuClicked(menuId, beforeEvent, out cancelEvent, out cancelMessage);
            }
        }

        /// <summary>
        /// Runs the right click.
        /// </summary>
        /// <param name="contextMenuInfo">The context menu info.</param>
        /// <param name="bubbleEvent">if set to <c>true</c> [bubble event].</param>
        private void RunRightClick(ref ContextMenuInfo contextMenuInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (Forms.ContainsKey(Application.Forms.ActiveForm.UniqueID))
            {
                IForm form = Forms[Application.Forms.ActiveForm.UniqueID];
                form.RightClicked(ref contextMenuInfo, out bubbleEvent);
            }
        }

        /// <summary>
        /// Applies the add on menu changes.
        /// </summary>
        private void ApplyAddOnMenuChanges()
        {
            try
            {
                var documentXml = new XmlDocument();
                string filePath = this.Location + "MenusChanges.xml";

                // load the content of the XML File
                documentXml.Load(filePath);

                // load the form to the SBO application in one batch
                string tmpStr = documentXml.InnerXml.Replace("$AddOnLocation$", this.Location);
                this.Application.LoadBatchActions(ref tmpStr);
            }
            catch { }
        }

        /// <summary>
        /// Applies the global menu changes.
        /// </summary>
        private void ApplyGlobalMenuChanges()
        {
            try
            {
                // Global Menus - in common to multiple addons
                var documentXml = new XmlDocument();
                string filePath = this.Location + "MenusChangesGlobal.xml";

                // load the content of the XML File
                documentXml.Load(filePath);

                // load the form to the SBO application in one batch
                string tmpStr = documentXml.InnerXml.Replace("$AddOnLocation$", this.Location);
                this.Application.LoadBatchActions(ref tmpStr);
            }
            catch { }
        }

        /// <summary>
        /// Load addon menu changes.
        /// </summary>
        public void LoadMenus()
        {
            this.ApplyGlobalMenuChanges();
            this.ApplyAddOnMenuChanges();
        }

        /// <summary>
        /// Handles the exception.
        /// </summary>
        /// <param name="ex">The ex.</param>
        public virtual void HandleException(Exception ex)
        {
            this.StatusBar.ErrorMessage = "Trapped Error: " + ex.Message;
            LogHandler.LogError(ex);
            // Raise the Exception handled event
        }

        /// <summary>
        /// Handles the exception.
        /// </summary>
        /// <param name="ex">The exception.</param>
        /// <param name="message">The message.</param>
        public void HandleException(Exception ex, string message)
        {
            this.StatusBar.ErrorMessage = "Trapped Error: " + ex.Message;
            LogHandler.LogError(ex, message);
            // Raise the Exception handled event
        }


        /// <summary>
        /// Exports the specified form.
        /// </summary>
        /// <param name="application">The application.</param>
        /// <param name="formUID">The form UID.</param>
        /// <param name="location">The location.</param>
        /// <returns>True is form is exported successully.</returns>
        public bool ExportForm(string formUID, string location)
        {
            try
            {
                SAPbouiCOM.Form form = this.Application.Forms.Item(formUID);
                string path = string.Format("{0}{1}.srf", location, form.Title.Replace("/", string.Empty).Replace(" ", string.Empty));

                // Create the file.
                using (FileStream fs = File.Create(path))
                {
                    string xmlForm = form.GetAsXML();

                    // Remove form id and type
                    string pattern = "appformnumber=\"[^\"]+\" ";
                    xmlForm = Regex.Replace(xmlForm, pattern, "appformnumber=\"1000001\" ");
                    pattern = "FormType=\"[^\"]+\" ";
                    xmlForm = Regex.Replace(xmlForm, pattern, "FormType=\"1000001\" ");
                    pattern = "\" uid=\"[^\"]+\" ";
                    xmlForm = Regex.Replace(xmlForm, pattern, "\" uid=\"MCI000001\" ");

                    byte[] info = new System.Text.UTF8Encoding(true).GetBytes(xmlForm);

                    fs.Write(info, 0, info.Length);                                     // Add some information to the file.
                    fs.Close();
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Opens the folder.
        /// </summary>
        /// <param name="folderName">Name of the folder.</param>
        public void OpenFolder(string folderName)
        {
            try
            {
                this.OpenApplication("explorer.exe", folderName);
            }
            catch (Exception ex)
            {
                this.HandleException(ex);
            }
        }

        /// <summary>
        /// Opens the application.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="arguments">The arguments.</param>
        public void OpenApplication(string fileName, string arguments)
        {
            try
            {
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = fileName;
                proc.StartInfo.Arguments = arguments;
                proc.Start();
            }
            catch (Exception ex)
            {
                this.HandleException(ex);
            }

        }

        public string OpenFileDialog(string filters, string initialDirectory)
        {
            if (string.IsNullOrEmpty(filters))
            {
                filters = "All files (*.*)|*.*";
            }

            if (string.IsNullOrEmpty(initialDirectory))
            {
                initialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            }

            var openFileDialog = new FileSelector
            {
                Filter = filters,
                InitialDirectory = initialDirectory
            };

            var getFileThread = new Thread(openFileDialog.GetFileName);

            // threadGetExcelFile.ApartmentState = ApartmentState.STA;
            getFileThread.SetApartmentState(ApartmentState.STA);
            try
            {
                getFileThread.Start();

                // Wait for thread to get started}
                while (!getFileThread.IsAlive)
                {
                }

                // Wait a sec more
                Thread.Sleep(1);

                // Wait for thread to end
                getFileThread.Join();

                // Use file name as you will here
                return openFileDialog.FileName;
            }
            catch
            {
                return string.Empty;
            }
        }


        private void SetLogPath()
        {
            
            // Get the log path
            string globalLogPath = string.Empty;
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(this.Company, "LogPath", out globalLogPath);
            }
            catch
            {
            }
            finally
            {
                if (string.IsNullOrEmpty(globalLogPath))
                {
                    globalLogPath = new SystemFolder(this.Company).LogFolder;
                    ConfigurationHelper.GlobalConfiguration.Save(this.Company, "LogPath", globalLogPath);
                }
            }

            // Make sure the config path matches the global one if not reset the config one
            string logPath = string.Empty;
            try
            {
                ConfigurationHelper.Load("LogPath", out logPath);
            }
            catch
            {
            }

            if (logPath != globalLogPath + _company.UserName + "\\")
            {
                ConfigurationHelper.Save("LogPath", globalLogPath + _company.UserName + "\\");
            }
        }
        #endregion Methods
    }
}