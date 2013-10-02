// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Form.cs" company="B1C Canada Inc.">
//   B1C Canada Inc.
// </copyright>
// <author>Frank Alunni</author>
// <summary>
//
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.UI.Windows
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml;
    using Model.EventArguments;
    using SAPbouiCOM;
    using B1C.Utility.Helpers;
    using B1C.SAP.DI.BusinessAdapters.Administration;
    using B1C.SAP.DI.BusinessAdapters;
    using B1C.SAP.DI.Helpers;

    #region Delegates

    /// <summary>
    /// Handles Item Click
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterItemClickHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles Item Pressed
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterItemPressedHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles Before validate
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeValidateHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles After validate
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterValidateHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles After Key Pressed
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterKeyDownHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles Before Key Pressed
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeKeyDownHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles Item Click
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void ButtonClickHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles Item double Click
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void DoubleClickHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles Folder Click
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void FolderClickHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles Drop down button Click
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void ChooseFromListHandler(object sender, ChooseFromListEventArgs e);

    /// <summary>
    /// Handles before form load
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeFormLoadHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles after form load
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterFormLoadHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles after arguments are set
    /// </summary>
    /// <param name="sender">The event sender.</param>
    public delegate void AfterArgumentsSetHandler(object sender);

    /// <summary>
    /// Handles before form close
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeFormCloseHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles the Before Add
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeAddHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles the Before Update
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeUpdateHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles the Before Save (Add/Update)
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeSaveHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles the Before data load
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeDataLoadHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles the After Add
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterAddHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles the After Update
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterUpdateHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles the After Save (Add/Update)
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterSaveHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles the After Data Load
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterDataLoadHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles after form resize
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void FormResizeHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles before add clicked
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeAddClickHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles before Right clicked
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeRightClickHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles After Right clicked
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterRightClickHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles before toolbar clicked
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeToolbarClickHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles After toolbar clicked
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterToolbarClickHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles before Menu clicked
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeMenuClickHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles After Menu clicked
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterMenuClickHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// Handles before find clicked
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeFindClickHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles after add click
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterAddClickHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Handles after find click
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterFindClickHandler(object sender, FormEventArgs e);

    /// <summary>
    /// Raised when the form mode change
    /// </summary>
    /// <param name="sender">The event sender.</param>
    public delegate void ModeChangedHandler(object sender);

    /// <summary>
    /// The Disposed Event
    /// </summary>
    /// <param name="sender">The event sender.</param>
    public delegate void DisposeHandler(object sender);

    /// <summary>
    /// The Combo select event
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void BeforeComboSelectHandler(object sender, ItemEventArgs e);

    /// <summary>
    /// The Combo select event
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">The event argument class.</param>
    public delegate void AfterComboSelectHandler(object sender, ItemEventArgs e);

    #endregion Delegates

    /// <summary>
    /// Instance of the Form class.
    /// </summary>
    /// <author>Frank Alunni</author>
    /// <remarks>The Form adapter transforms and simplifies the usage of a SAPbouiCOM.Form. This is the only adapter that does not inherit from the BaseAdapter because of the different nature of the object itself.</remarks>
    public abstract class Form : IDisposable, Model.Interfaces.IForm
    {
        #region Fields

        /// <summary>
        /// The open this.arguments
        /// </summary>
        private Dictionary<string, object> _arguments;

        /// <summary>
        /// True if the changes have been applied;
        /// </summary>
        private bool _changesApplied;

        /// <summary>
        /// True if the screen need to be resized
        /// </summary>
        private bool _needsResize;

        /// <summary>
        /// True if the screen need to be reloaded
        /// </summary>
        private bool _needsReload;

        /// <summary>
        /// List of context menu to hide
        /// </summary>
        private List<string> _contextMenuHidden;

        /// <summary>
        /// List of context menu to remove
        /// </summary>
        private List<string> _contextMenuDisabled;

        /// <summary>
        /// The field name for the document field
        /// </summary>
        private string _findDocumentFieldName = string.Empty;

        /// <summary>
        /// The document code to be found
        /// </summary>
        private string _findDocumentNumber = string.Empty;

        #endregion Fields

        #region Constuctors

        /// <summary>
        /// Initializes a new instance of the <see cref="Form"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        protected Form(AddOn addOn)
            : this(addOn, addOn.Application.Forms.ActiveForm)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Form"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="form">The SAP form.</param>
        protected Form(AddOn addOn, SAPbouiCOM.Form form)
        {
            this.LoadFormClass(addOn, form);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Form"/> class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        protected Form(AddOn addOn, string xmlPath, string formType)
            : this(addOn, xmlPath, formType, string.Empty)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Form"/> class.
        /// </summary>
        /// <param name="addOn">The running add on.</param>
        /// <param name="xmlPath">The XML path.</param>
        /// <param name="formType">Type of the form.</param>
        /// <param name="objectType">Type of the object.</param>
        protected Form(AddOn addOn, string xmlPath, string formType, string objectType)
        {
            this.AddOn = addOn;

            this.AddOn.FormCount += 1;

            this.FormId = string.Format("{0}_{1}{2}", this.AddOn.Namespace, this.AddOn.AddonId, this.AddOn.FormCount);
            this.FormTypeId = string.Format("{0}_{1}{2}", this.AddOn.Namespace, this.AddOn.AddonId, formType);

            // If form not already loaded get it from XML
            try
            {
                string color;
                try
                {
                    var xmlmain = new XmlDocument();
                    xmlmain.LoadXml(this.AddOn.Application.Forms.ActiveForm.GetAsXML());
                    var node = xmlmain.SelectSingleNode("Application/forms/action/form/@color");
                    color = node != null
                                ? (node.Value ?? string.Empty)
                                : string.Empty;
                }
                catch
                {
                    color = string.Empty;
                }

                var xmlReader = new XmlTextReader(xmlPath);
                var xmldoc = new XmlDocument();
                xmldoc.Load(xmlReader);
                xmlReader.Close();

                var creationParam = (FormCreationParams)this.AddOn.Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams);

                creationParam.FormType = this.FormTypeId;
                creationParam.UniqueID = this.FormId;

                if (!string.IsNullOrEmpty(objectType))
                {
                    creationParam.ObjectType = objectType;
                }

                if (!string.IsNullOrEmpty(color))
                {
                    var node = xmldoc.SelectSingleNode("Application/forms/action/form/@color");
                    if (node != null)
                    {
                        node.Value = color;
                    }
                }

                creationParam.XmlData = xmldoc.InnerXml;

                // Create form from the creation parameter
                this.SapForm = this.AddOn.Application.Forms.AddEx(creationParam);

                // Do not allow to remove records
                this.SapForm.EnableMenu("1283", false);

                // Do not allow to restore records
                this.SapForm.EnableMenu("1285", false);

                // Do not allow to duplicate records
                this.SapForm.EnableMenu("1287", false);

                // Do not allow to Add Row
                this.SapForm.EnableMenu("1292", false);

                // Do not allow to Delete Row
                this.SapForm.EnableMenu("1293", false);

                // Do not allow to Duplicate Row
                this.SapForm.EnableMenu("1294", false);

                // Do not allow to Close Row
                this.SapForm.EnableMenu("1299", false);

                this.Initialize();
                this._needsResize = true;
                this._needsReload = true;
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("Internal error (-10)"))
                {
                    this.AddOn.HandleException(ex);
                }
            }
        }

        #endregion Constuctors

        #region Events

        /// <summary>
        /// Occurs when [item click handler].
        /// </summary>
        public event AfterItemClickHandler AfterItemClick;

        /// <summary>
        /// Occurs when [item pressed handler].
        /// </summary>
        public event AfterItemPressedHandler AfterItemPressed;

        /// <summary>
        /// Occurs when [Before validate].
        /// </summary>
        public event BeforeValidateHandler BeforeValidate;

        /// <summary>
        /// Occurs when [after validate].
        /// </summary>
        public event AfterValidateHandler AfterValidate;

        /// <summary>
        /// Occurs when [after key pressed].
        /// </summary>
        public event AfterKeyDownHandler AfterKeyDown;

        /// <summary>
        /// Occurs when [Before key pressed].
        /// </summary>
        public event BeforeKeyDownHandler BeforeKeyDown;

        /// <summary>
        /// Occurs when [item combo select].
        /// </summary>
        public event BeforeComboSelectHandler BeforeComboSelect;

        /// <summary>
        /// Occurs when [item combo select].
        /// </summary>
        public event AfterComboSelectHandler AfterComboSelect;

        /// <summary>
        /// Occurs when [item click handler].
        /// </summary>
        public event ButtonClickHandler ButtonClick;

        /// <summary>
        /// Occurs when [item double click handler].
        /// </summary>
        public event DoubleClickHandler DoubleClick;

        /// <summary>
        /// Occurs when [folder click].
        /// </summary>
        public event FolderClickHandler FolderClick;

        /// <summary>
        /// Occurs when [drop down button click].
        /// </summary>
        public event ChooseFromListHandler ChooseFromList;

        /// <summary>
        /// Occurs when [before form load].
        /// </summary>
        public event BeforeFormLoadHandler BeforeFormLoad;

        /// <summary>
        /// Occurs when [after form load].
        /// </summary>
        public event AfterFormLoadHandler AfterFormLoad;

        /// <summary>
        /// Occurs when [after form load].
        /// </summary>
        public event AfterArgumentsSetHandler AfterArgumentsSet;

        /// <summary>
        /// Occurs when [before form close].
        /// </summary>
        public event BeforeFormCloseHandler BeforeFormClose;

        /// <summary>
        /// Occurs when [before add].
        /// </summary>
        public event BeforeAddHandler BeforeAdd;

        /// <summary>
        /// Occurs when [before update].
        /// </summary>
        public event BeforeUpdateHandler BeforeUpdate;

        /// <summary>
        /// Occurs when [before save].
        /// </summary>
        public event BeforeSaveHandler BeforeSave;

        /// <summary>
        /// Occurs when [before data load].
        /// </summary>
        public event BeforeDataLoadHandler BeforeDataLoad;

        /// <summary>
        /// Occurs when [after add].
        /// </summary>
        public event AfterAddHandler AfterAdd;

        /// <summary>
        /// Occurs when [after update].
        /// </summary>
        public event AfterUpdateHandler AfterUpdate;

        /// <summary>
        /// Occurs when [after save].
        /// </summary>
        public event AfterSaveHandler AfterSave;

        /// <summary>
        /// Occurs when [after data load].
        /// </summary>
        public event AfterDataLoadHandler AfterDataLoad;

        /// <summary>
        /// Occurs when [form resize].
        /// </summary>
        public event FormResizeHandler FormResize;

        /// <summary>
        /// Occurs when [before add Click].
        /// </summary>
        public event BeforeAddClickHandler BeforeAddClick;

        /// <summary>
        /// Occurs when [before Right Click].
        /// </summary>
        public event BeforeRightClickHandler BeforeRightClick;

        /// <summary>
        /// Occurs when [After Right Click].
        /// </summary>
        public event AfterRightClickHandler AfterRightClick;

        /// <summary>
        /// Occurs when [before toolbar Click].
        /// </summary>
        public event BeforeToolbarClickHandler BeforeToolbarClick;

        /// <summary>
        /// Occurs when [After toolbar Click].
        /// </summary>
        public event AfterToolbarClickHandler AfterToolbarClick;

        /// <summary>
        /// Occurs when [before Menu Click].
        /// </summary>
        public event BeforeMenuClickHandler BeforeMenuClick;

        /// <summary>
        /// Occurs when [After Menu Click].
        /// </summary>
        public event AfterMenuClickHandler AfterMenuClick;

        /// <summary>
        /// Occurs when [before find Click].
        /// </summary>
        public event BeforeFindClickHandler BeforeFindClick;

        /// <summary>
        /// Occurs when [after add click].
        /// </summary>
        public event AfterAddClickHandler AfterAddClick;

        /// <summary>
        /// Occurs when [after find click].
        /// </summary>
        public event AfterFindClickHandler AfterFindClick;

        /// <summary>
        /// Occurs when [disposed].
        /// </summary>
        public event ModeChangedHandler ModeChanged;

        /// <summary>
        /// Occurs when [disposed].
        /// </summary>
        public event DisposeHandler Disposed;

        #endregion Events

        #region Properties

        /// <summary>
        /// Gets the SAP form.
        /// </summary>
        /// <value>The sap form.</value>
        public SAPbouiCOM.Form SapForm
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the application.
        /// </summary>
        /// <value>The application.</value>
        public AddOn AddOn
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the form id.
        /// </summary>
        /// <value>The form id.</value>
        public string FormId
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the form type id.
        /// </summary>
        /// <value>The form id.</value>
        public string FormTypeId
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the data source.
        /// </summary>
        /// <value>The data source.</value>
        public Data.DataSource DataSource { get; private set; }

        /// <summary>
        /// Gets the control manager.
        /// </summary>
        /// <value>The control manager.</value>
        public ControlManager ControlManager { get; private set; }

        /// <summary>
        /// Gets or sets the mode.
        /// </summary>
        /// <value>The form mode.</value>
        public BoFormMode Mode
        {
            get
            {
                return this.SapForm.Mode;
            }

            set
            {
                this.SapForm.Mode = value;
                this.OnModeChanged();
            }
        }

        /// <summary>
        /// Gets the name of the table.
        /// </summary>
        /// <value>The name of the table.</value>
        public string TableName
        {
            get
            {
                try
                {
                    return ((DBDataSourceClass)(this.SapForm.DataSources.DBDataSources.Item(0))).TableName;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// Gets or sets the this.arguments.
        /// </summary>
        /// <value>The this.arguments.</value>
        public Dictionary<string, object> Arguments
        {
            get { return this._arguments ?? (this._arguments = new Dictionary<string, object>()); }

            set
            {
                this._arguments = value;

                this.Freeze(true);
                try
                {
                    this.OnAfterArgumentsSet();
                }
                finally
                {
                    this.Freeze(false);
                }
            }
        }

        /// <summary>
        /// Gets or sets the hide context menu.
        /// </summary>
        /// <value>The hide context menu.</value>
        public List<String> ContextMenuHidden
        {
            get
            {
                if (_contextMenuHidden == null)
                {
                    this.AddOn.AddFilter(BoEventTypes.et_RIGHT_CLICK, this.FormTypeId);
                    _contextMenuHidden = new List<string>();
                }

                return _contextMenuHidden;
            }
        }

        /// <summary>
        /// Gets or sets the removed context menu.
        /// </summary>
        /// <value>The hide context menu.</value>
        public List<String> ContextMenuDisabled
        {
            get
            {
                if (_contextMenuDisabled == null)
                {
                    this.AddOn.AddFilter(BoEventTypes.et_RIGHT_CLICK, this.FormTypeId);
                    _contextMenuDisabled = new List<string>();
                }

                return _contextMenuDisabled;
            }
        }

        #endregion Properties

        #region Method

        /// <summary>
        /// Determines whether the specified subpermission is authorized.
        /// </summary>
        /// <param name="subpermission">The subpermission.</param>
        /// <returns>
        /// 	<c>true</c> if the specified subpermission is authorized; otherwise, <c>false</c>.
        /// </returns>
        public bool IsAuthorized(string subpermission)
        {
            string formType = this.SapForm.TypeEx;
            var currentUser = UserAdapter.Current(this.AddOn.Company);

            if (currentUser.Superuser)
            {
                return true;
            }

            // If permission does not exists users are authorized
            using (var rs = new RecordsetAdapter(this.AddOn.Company, string.Format("SELECT * FROM UPT1 Where PermId = '{0}_{1}'", formType, subpermission)))
            {
                if (rs.EoF)
                {
                    return true;
                }
            }

            // If permission exists users must be authorized
            string query = string.Format("SELECT * FROM USR3 Where Permission = 'F' AND PermId = '{0}_{1}' AND UserLink = {2}", formType, subpermission, currentUser.InternalKey);
            using (var rs = new RecordsetAdapter(this.AddOn.Company, query))
            {
                return !rs.EoF;
            }
        }

        /// <summary>
        /// Loads the form class.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <param name="form">The form.</param>
        private void LoadFormClass(AddOn addOn, SAPbouiCOM.Form form)
        {
            try
            {
                this.AddOn = addOn;
                this.SapForm = form;
                this.FormId = this.SapForm.UniqueID;
                this.FormTypeId = this.SapForm.TypeEx;
                this.SapForm.Freeze(true);
                this.Initialize();
                this.ApplyChanges();
            }
            catch (Exception ex)
            {
                // If an error occurr, log it and fails the method
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Finds the document.
        /// </summary>
        /// <param name="docNumFieldName">Name of the doc num field.</param>
        /// <param name="docNum">The doc num.</param>
        public void FindDocument(string docNumFieldName, string docNum)
        {
            this._findDocumentFieldName = docNumFieldName;
            this._findDocumentNumber = docNum;
            this.SelectDocument();
        }

        /// <summary>
        /// Popups the message.
        /// </summary>
        /// <param name="message">The message.</param>
        protected void PopupMessage(string message)
        {
            try
            {
                var popUp = this.AddOn.Application.Forms.Add("popUp", BoFormTypes.ft_Fixed, 99999);
                popUp.DataSources.UserDataSources.Add("messageDS", BoDataType.dt_LONG_TEXT, 64000);
                var messageEdt = popUp.Items.Add("message", BoFormItemTypes.it_EXTEDIT);
                popUp.Top = 50;
                popUp.Left = 50;
                popUp.Width = 600;
                popUp.Height = 250;
                messageEdt.Top = 5;
                messageEdt.Left = 5;
                messageEdt.Width = 585;
                messageEdt.Height = 205;
                var edit = (EditText)messageEdt.Specific;
                edit.DataBind.SetBound(true, string.Empty, "messageDS");

                popUp.DataSources.UserDataSources.Item("messageDS").Value = message;

                popUp.Visible = true;
            }
            catch
            {
                var popUp = this.AddOn.Application.Forms.Item("popUp");
                popUp.DataSources.UserDataSources.Item("messageDS").Value = message;
                popUp.Visible = true;
            }
        }

        /// <summary>
        /// Closes this instance.
        /// </summary>
        /// <returns>True if form closes successfully.</returns>
        public bool Close()
        {
            try
            {
                this.SapForm.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Fields the specified field name.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>The field requested</returns>
        public Data.Field Field(string fieldName)
        {
            return new Data.Field(this, fieldName);
        }

        /// <summary>
        /// Fields the specified table name.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>The field requested</returns>
        public Data.Field Field(string tableName, string fieldName)
        {
            return new Data.Field(this, tableName, fieldName);
        }

        /// <summary>
        /// Refreshes this instance.
        /// </summary>
        public void Refresh()
        {
            if (this._needsReload)
            {
                this._needsReload = false;
                this.OnAfterFormLoad(new FormEventArgs(this.FormTypeId, this.FormId, string.Empty));
            }

            if (this._needsResize)
            {
                this._needsResize = false;
                this.OnFormResize(new FormEventArgs(this.FormTypeId, this.FormId, string.Empty));
            }
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public virtual void Dispose()
        {
            Debug.Print(string.Format("Disposing form {0} of form type {1}", this.FormId, this.FormTypeId));

            this.AddOn.Application.ItemEvent -= this.ApplicationItemEvent;
            this.AddOn.Application.FormDataEvent -= this.ApplicationFormDataEvent;
            this.DataSource = null;
            this.ControlManager = null;
            this.SapForm = null;
            this.AddOn = null;

            this.OnDispose();
        }

        /// <summary>
        /// Determines whether [is sub table editable].
        /// </summary>
        /// <returns>True if the sub tables can be edited</returns>
        public bool IsSubTableEditable()
        {
            try
            {
                switch (this.SapForm.Mode)
                {
                    case BoFormMode.fm_FIND_MODE:
                        this.AddOn.StatusBar.ErrorMessage = "This operation is not permitted if the form is in 'Find' mode.";
                        return false;
                    case BoFormMode.fm_ADD_MODE:
                        this.AddOn.StatusBar.ErrorMessage = "This operation is not permitted if the form is in 'Add' mode.";
                        return false;
                    case BoFormMode.fm_UPDATE_MODE:
                        this.AddOn.StatusBar.ErrorMessage = "This operation is not permitted if the form is in 'Update' mode.";
                        return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
                return false;
            }
        }

        /// <summary>
        /// Freezes the specified is frozen.
        /// </summary>
        /// <param name="isFrozen">if set to <c>true</c> [is frozen].</param>
        public void Freeze(bool isFrozen)
        {
            try
            {
                this.SapForm.Freeze(isFrozen);
            }
            catch
            {
            }
        }

        /// <summary>
        /// Determines whether [is custom form].
        /// </summary>
        /// <returns>
        /// <c>true</c> if [is custom form]; otherwise, <c>false</c>.
        /// </returns>
        public bool IsCustomForm()
        {
            return this.FormId.StartsWith(string.Format("{0}_{1}", this.AddOn.Namespace, this.AddOn.AddonId));
        }

        /// <summary>
        /// Opens the document.
        /// </summary>
        /// <param name="documentType">Type of the document.</param>
        /// <param name="documentEntry">The document entry.</param>
        public void OpenDocument(BoLinkedObject documentType, int documentEntry)
        {
            try
            {
                this.Freeze(true);

                Item editItem;

                // Retrieve the Edit Item
                try
                {
                    editItem = this.SapForm.Items.Item("openDcEdit");
                }
                catch
                {
                    editItem = this.SapForm.Items.Add("openDcEdit", BoFormItemTypes.it_EDIT);
                    editItem.Top = -100;
                    editItem.AffectsFormMode = false;
                }

                // Set the edit text value
                ((EditText)editItem.Specific).Value = documentEntry.ToString();

                Item linkItem;

                // Retrieve the linked button item
                try
                {
                    linkItem = this.SapForm.Items.Item("openDcLink");
                }
                catch
                {
                    linkItem = this.SapForm.Items.Add("openDcLink", BoFormItemTypes.it_LINKED_BUTTON);
                    linkItem.Top = -100;
                    linkItem.LinkTo = "openDcEdit";
                    linkItem.AffectsFormMode = false;
                }

                ((LinkedButton)linkItem.Specific).LinkedObject = documentType;

                this.SapForm.Items.Item("openDcLink").Click(BoCellClickType.ct_Linked);
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
            finally
            {
                this.Freeze(false);
            }
        }

        /// <summary>
        /// Applies the filters.
        /// </summary>
        public void ApplyFilters()
        {
            if ((this.AfterItemClick != null) || (this.ButtonClick != null) || (this.FolderClick != null))
            {
                this.AddOn.AddFilter(BoEventTypes.et_CLICK, this.FormTypeId);
            }

            if (this.FormResize != null)
            {
                this.AddOn.AddFilter(BoEventTypes.et_FORM_RESIZE, this.FormTypeId);
            }

            if (this.AfterComboSelect != null)
            {
                this.AddOn.AddFilter(BoEventTypes.et_COMBO_SELECT, this.FormTypeId);
            }

            // Add The form Data filters
            if ((this.BeforeDataLoad != null) || (this.AfterDataLoad != null))
            {
                this.AddOn.AddFilter(BoEventTypes.et_FORM_DATA_LOAD, this.FormTypeId);
            }
            if ((this.AfterAdd != null) || (this.AfterSave != null) || (this.BeforeAdd != null) || (this.BeforeSave != null))
            {
                this.AddOn.AddFilter(BoEventTypes.et_FORM_DATA_ADD, this.FormTypeId);
            }

            if ((this.AfterUpdate != null) || (this.AfterSave != null) || (this.BeforeUpdate != null) || (this.BeforeSave != null))
            {
                this.AddOn.AddFilter(BoEventTypes.et_FORM_DATA_UPDATE, this.FormTypeId);
            }

            if (this.AfterItemPressed != null)
            {
                this.AddOn.AddFilter(BoEventTypes.et_ITEM_PRESSED, this.FormTypeId);
            }

            if (this.BeforeFormClose != null)
            {
                this.AddOn.AddFilter(BoEventTypes.et_FORM_CLOSE, this.FormTypeId);
            }

            if (this.DoubleClick != null)
            {
                this.AddOn.AddFilter(BoEventTypes.et_DOUBLE_CLICK, this.FormTypeId);
            }

            if (this.AfterValidate != null)
            {
                this.AddOn.AddFilter(BoEventTypes.et_VALIDATE, this.FormTypeId);
            }

            if (this.ChooseFromList != null)
            {
                this.AddOn.AddFilter(BoEventTypes.et_CHOOSE_FROM_LIST, this.FormTypeId);
            }

            if ((this.AfterKeyDown != null) || (this.BeforeKeyDown != null))
            {
                this.AddOn.AddFilter(BoEventTypes.et_KEY_DOWN, this.FormTypeId);
            }

            if ((this.AfterToolbarClick != null) || (this.BeforeToolbarClick != null)
             || (this.AfterMenuClick != null) || (this.BeforeMenuClick != null))
            {
                this.AddOn.AddFilter(BoEventTypes.et_MENU_CLICK, string.Empty);
            }

            return;
        }

        /// <summary>
        /// Tools the bar clicked.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        /// <param name="beforeEvent">if set to <c>true</c> [before event].</param>
        /// <param name="cancelEvent">if set to <c>true</c> [cancel event].</param>
        /// <param name="cancelMessage">The cancel message.</param>
        public void ToolBarClicked(string buttonId, bool beforeEvent, out bool cancelEvent, out string cancelMessage)
        {
            var arguments = new FormEventArgs(this.FormTypeId, this.FormId, string.Empty);
            var toolbarArguments = new ItemEventArgs(this.FormTypeId, 0, this.FormId, buttonId, string.Empty, -1);

            if (beforeEvent)
            {
                this.OnBeforeToolbarClick(toolbarArguments);
                if (!toolbarArguments.CancelEvent)
                {
                    switch (buttonId)
                    {
                        case Constants.Toolbar.Add:
                            this.OnBeforeAddClick(arguments);
                            break;
                        case Constants.Toolbar.Find:
                            this.OnBeforeFindClick(arguments);
                            break;
                    }
                }
            }
            else
            {
                this.OnAfterToolbarClick(toolbarArguments);
                if (!toolbarArguments.CancelEvent)
                {
                    switch (buttonId)
                    {
                        case Constants.Toolbar.Add:
                            this.OnAfterAddClick(arguments);
                            break;
                        case Constants.Toolbar.Find:
                            this.OnAfterFindClick(arguments);
                            break;
                    }
                }
            }

            if (!toolbarArguments.CancelEvent)
            {
                cancelEvent = arguments.CancelEvent;
                cancelMessage = arguments.CancelMessage;
            }
            else
            {
                cancelEvent = toolbarArguments.CancelEvent;
                cancelMessage = toolbarArguments.CancelMessage;
            }
        }


        /// <summary>
        /// Menu was clicked.
        /// </summary>
        /// <param name="buttonId">The button id.</param>
        /// <param name="beforeEvent">if set to <c>true</c> [before event].</param>
        /// <param name="cancelEvent">if set to <c>true</c> [cancel event].</param>
        /// <param name="cancelMessage">The cancel message.</param>
        public void MenuClicked(string buttonId, bool beforeEvent, out bool cancelEvent, out string cancelMessage)
        {
            var arguments = new FormEventArgs(this.FormTypeId, this.FormId, string.Empty);
            var toolbarArguments = new ItemEventArgs(this.FormTypeId, 0, this.FormId, buttonId, string.Empty, -1);

            if (beforeEvent)
            {
                this.OnBeforeMenuClick(toolbarArguments);
            }
            else
            {
                this.OnAfterMenuClick(toolbarArguments);
            }

            if (!toolbarArguments.CancelEvent)
            {
                cancelEvent = arguments.CancelEvent;
                cancelMessage = arguments.CancelMessage;
            }
            else
            {
                cancelEvent = toolbarArguments.CancelEvent;
                cancelMessage = toolbarArguments.CancelMessage;
            }
        }


        /// <summary>
        /// Rights the clicked.
        /// </summary>
        /// <param name="contextMenuInfo">The context menu info.</param>
        /// <param name="bubbleEvent">if set to <c>true</c> [bubble event].</param>
        public void RightClicked(ref ContextMenuInfo contextMenuInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            var toolbarArguments = new ItemEventArgs(this.FormTypeId, 0, this.FormId, contextMenuInfo.ItemUID, contextMenuInfo.ColUID, contextMenuInfo.Row);

            if (contextMenuInfo.BeforeAction)
            {
                this.OnBeforeRightClick(toolbarArguments);

                // Hide menus
                if (this._contextMenuHidden != null)
                {
                    foreach (var hiddenMenu in this._contextMenuHidden)
                    {
                        contextMenuInfo.RemoveFromContent(hiddenMenu);
                    }
                }

                // Disabled Menus
                if (this._contextMenuDisabled != null)
                {
                    foreach (var disabledMenu in this._contextMenuDisabled)
                    {
                        this.AddOn.Application.Menus.Item(disabledMenu).Enabled = false;
                    }
                }
            }
            else
            {
                this.OnAfterRightClick(toolbarArguments);
            }
        }

        /// <summary>
        /// Applications the form data event.
        /// </summary>
        /// <param name="businessObjectInfo">The business object info.</param>
        /// <param name="bubbleEvent">if set to <c>true</c> [bubble event].</param>
        internal void ApplicationFormDataEvent(ref BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (this.FormId != businessObjectInfo.FormUID)
            {
                return;
            }

            var arguments = new FormEventArgs(this.FormTypeId, this.FormId, businessObjectInfo.ObjectKey);

            if (businessObjectInfo.BeforeAction)
            {
                switch (businessObjectInfo.EventType)
                {
                    case BoEventTypes.et_FORM_DATA_LOAD:
                        this.OnBeforeDataLoad(arguments);
                        break;
                    case BoEventTypes.et_FORM_DATA_ADD:
                        this.OnBeforeSave(arguments);
                        this.OnBeforeAdd(arguments);
                        break;
                    case BoEventTypes.et_FORM_DATA_UPDATE:
                        this.OnBeforeSave(arguments);
                        this.OnBeforeUpdate(arguments);
                        break;
                }
            }
            else
            {
                arguments.IsSuccessful = businessObjectInfo.ActionSuccess;

                switch (businessObjectInfo.EventType)
                {
                    case BoEventTypes.et_FORM_DATA_LOAD:
                        this.OnAfterDataLoad(arguments);
                        break;
                    case BoEventTypes.et_FORM_DATA_ADD:
                        this.OnAfterSave(arguments);
                        this.OnAfterAdd(arguments);
                        break;
                    case BoEventTypes.et_FORM_DATA_UPDATE:
                        this.OnAfterSave(arguments);
                        this.OnAfterUpdate(arguments);
                        break;
                }
            }

            bubbleEvent = !arguments.CancelEvent;

            if (arguments.CancelEvent && !string.IsNullOrEmpty(arguments.CancelMessage))
            {
                this.AddOn.ActionCancelledMessage = arguments.CancelMessage;
            }
        }

        /// <summary>
        /// Application_s the item event.
        /// </summary>
        /// <param name="formId">
        /// The form UID.
        /// </param>
        /// <param name="itemEvent">
        /// The item event.
        /// </param>
        /// <param name="bubbleEvent">
        /// if set to <c>true</c> [bubble event].
        /// </param>
        internal void ApplicationItemEvent(string formId, ref ItemEvent itemEvent, out bool bubbleEvent)
        {
            bubbleEvent = true;
            string task = string.Empty;

            System.Diagnostics.Debug.WriteLine(itemEvent.EventType);

            try
            {
                task = "Checking Form Id";
                if (this.FormId != formId)
                {
                    return;
                }

                task = "Checking Item";
                Item item = null;
                if (!string.IsNullOrEmpty(itemEvent.ItemUID))
                {
                    item = this.SapForm.Items.Item(itemEvent.ItemUID);
                    if (item == null)
                    {
                        return;
                    }

                    task = "Checking Item enabled";
                    if (item.Enabled == false)
                    {
                        return;
                    }
                }

                // Only now you can raise the event
                task = "Checking After Event Type";
                System.Diagnostics.Debug.WriteLine(itemEvent.EventType);
                if (itemEvent.BeforeAction)
                {
                    switch (itemEvent.EventType)
                    {
                        case BoEventTypes.et_FORM_LOAD:
                            this.OnBeforeFormLoad(new FormEventArgs(itemEvent.FormTypeEx, itemEvent.FormUID, string.Empty));
                            break;
                        case BoEventTypes.et_FORM_CLOSE:
                            this.OnBeforeFormClose(new FormEventArgs(itemEvent.FormTypeEx, itemEvent.FormUID, string.Empty));
                            break;
                        case BoEventTypes.et_KEY_DOWN:
                            this.OnBeforeKeyDown(new ItemEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, itemEvent.CharPressed));
                            break;
                        case BoEventTypes.et_COMBO_SELECT:
                            task = "Rasing Combo select";
                            var arguments = new ItemEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, itemEvent.ColUID, itemEvent.Row);
                            if (itemEvent.ItemUID != null)
                            {
                                this.OnBeforeComboSelect(arguments);
                            }

                            bubbleEvent = !arguments.CancelEvent;
                            if (arguments.CancelEvent && !string.IsNullOrEmpty(arguments.CancelMessage))
                            {
                                this.AddOn.ActionCancelledMessage = arguments.CancelMessage;
                            }

                            break;
                        case BoEventTypes.et_VALIDATE:
                            arguments = new ItemEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, itemEvent.ColUID, itemEvent.Row);
                            this.OnBeforeValidate(arguments);
                            bubbleEvent = !arguments.CancelEvent;
                            if (arguments.CancelEvent && !string.IsNullOrEmpty(arguments.CancelMessage))
                            {
                                this.AddOn.ActionCancelledMessage = arguments.CancelMessage;
                            }

                            break;
                    }
                }
                else
                {
                    switch (itemEvent.EventType)
                    {
                        case BoEventTypes.et_CHOOSE_FROM_LIST:
                            var chooseFromListEvent = (IChooseFromListEvent)itemEvent;
                            // Selected objects comes back as NULL when the 'Cancel button is clicked'

                            int selectedRow = itemEvent.Row;
                            
                            if (chooseFromListEvent.SelectedObjects != null)
                            {
                                this.OnChooseFromList(new ChooseFromListEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, chooseFromListEvent.SelectedObjects, itemEvent.Row, itemEvent.ColUID));
                            }

                            break;
                        case BoEventTypes.et_FORM_LOAD:
                            // this.SapForm.Freeze(true);
                            this.OnAfterFormLoad(new FormEventArgs(itemEvent.FormTypeEx, itemEvent.FormUID, string.Empty));
                            this.SapForm.Freeze(false);
                            break;
                        case BoEventTypes.et_FORM_RESIZE:
                            if (this.SapForm.State != BoFormStateEnum.fs_Minimized)
                            {
                                this.OnFormResize(new FormEventArgs(itemEvent.FormTypeEx, itemEvent.FormUID, string.Empty));
                            }
                            break;
                        case BoEventTypes.et_COMBO_SELECT:
                            task = "Rasing Combo select";
                            if (itemEvent.ItemUID != null)
                            {
                                this.OnAfterComboSelect(new ItemEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, itemEvent.ColUID, itemEvent.Row));
                            }

                            break;
                        case BoEventTypes.et_VALIDATE:
                            if (itemEvent.ItemChanged)
                            {
                                this.OnAfterValidate(new ItemEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, itemEvent.ColUID, itemEvent.Row));
                            }
                            break;
                        case BoEventTypes.et_DOUBLE_CLICK:
                            task = "Rasing Double Click Event";
                            this.OnDoubleClick(new ItemEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, itemEvent.ColUID, itemEvent.Row));
                            break;
                        case BoEventTypes.et_ITEM_PRESSED:
                            task = "Rasing Double Click Event";
                            this.OnAfterItemPressed(new ItemEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, itemEvent.ColUID, itemEvent.Row));
                            break;
                        case BoEventTypes.et_KEY_DOWN:
                            this.OnAfterKeyDown(new ItemEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, itemEvent.CharPressed));
                            break;
                        case BoEventTypes.et_CLICK:
                            if (item != null)
                            {
                                switch (item.Type)
                                {
                                    case BoFormItemTypes.it_FOLDER:
                                        task = "Rasing Button Event";
                                        this.OnFolderClick(new ItemEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, itemEvent.ColUID, itemEvent.Row));
                                        break;
                                    case BoFormItemTypes.it_BUTTON_COMBO:
                                    case BoFormItemTypes.it_BUTTON:
                                        task = "Rasing Button Event";
                                        this.OnButtonClick(new ItemEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, itemEvent.ColUID, itemEvent.Row));
                                        break;
                                    default:
                                        task = "Rasing Item Event";
                                        this.OnAfterItemClick(new ItemEventArgs(itemEvent.FormTypeEx, itemEvent.FormMode, itemEvent.FormUID, itemEvent.ItemUID, itemEvent.ColUID, itemEvent.Row));
                                        break;
                                }
                            }

                            break;
                        case BoEventTypes.et_FORM_UNLOAD:
                            this.Dispose();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex, task);
            }
        }

        /// <summary>
        /// Called when [dispose].
        /// </summary>
        internal void OnDispose()
        {
            try
            {
                if (this.Disposed != null)
                {
                    this.Disposed(this);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Initializes this instance.
        /// </summary>
        public virtual void Initialize()
        {
            Debug.Print(string.Format("Initializing form {0} of form type {1}", this.FormId, this.FormTypeId));
            this.DataSource = new Data.DataSource(this);
            this.ControlManager = new ControlManager(this);

            this.AddOn.Application.ItemEvent += this.ApplicationItemEvent;
            this.AddOn.Application.FormDataEvent += this.ApplicationFormDataEvent;
        }

        /// <summary>
        /// Called when [mode changed].
        /// </summary>
        protected void OnModeChanged()
        {
            try
            {
                if (this.ModeChanged != null)
                {
                    this.ModeChanged(this);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the AfterItemClick event.
        /// </summary>
        /// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterItemClick(ItemEventArgs e)
        {
            try
            {
                if (this.AfterItemClick != null)
                {
                    this.AfterItemClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the AfterItemPressed event.
        /// </summary>
        /// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterItemPressed(ItemEventArgs e)
        {
            try
            {
                if (this.AfterItemPressed != null)
                {
                    this.AfterItemPressed(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="BeforeValidate"/> event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeValidate(ItemEventArgs e)
        {
            try
            {
                if (this.BeforeValidate != null)
                {
                    this.BeforeValidate(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="AfterValidate"/> event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterValidate(ItemEventArgs e)
        {
            try
            {
                if (this.AfterValidate != null)
                {
                    this.AfterValidate(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="OnAfterKeyDown"/> event.
        /// </summary>
        /// <param name="e">The <see cref="Maritz.SAP.UI.Model.EventArguments.ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterKeyDown(ItemEventArgs e)
        {
            try
            {
                if (this.AfterKeyDown != null)
                {
                    this.AfterKeyDown(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="OnBeforeKeyDown"/> event.
        /// </summary>
        /// <param name="e">The <see cref="Maritz.SAP.UI.Model.EventArguments.ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeKeyDown(ItemEventArgs e)
        {
            try
            {
                if (this.BeforeKeyDown != null)
                {
                    this.BeforeKeyDown(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="ComboSelect"/> event.
        /// </summary>
        /// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeComboSelect(ItemEventArgs e)
        {
            try
            {
                if (this.BeforeComboSelect != null)
                {
                    this.BeforeComboSelect(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="ComboSelect"/> event.
        /// </summary>
        /// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterComboSelect(ItemEventArgs e)
        {
            try
            {
                if (this.AfterComboSelect != null)
                {
                    this.AfterComboSelect(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the AfterItemClick event.
        /// </summary>
        /// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnButtonClick(ItemEventArgs e)
        {
            try
            {
                if (this.ButtonClick != null)
                {
                    this.ButtonClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the AfterItemDoubleClick event.
        /// </summary>
        /// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnDoubleClick(ItemEventArgs e)
        {
            try
            {
                if (this.DoubleClick != null)
                {
                    this.DoubleClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the FolderClick event.
        /// </summary>
        /// <param name="e">The <see cref="ItemEventArgs"/> instance containing the event data.</param>
        protected virtual void OnFolderClick(ItemEventArgs e)
        {
            try
            {
                if (this.FolderClick != null)
                {
                    this.FolderClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the AfterItemClick event.
        /// </summary>
        /// <param name="e">The <see cref="ChooseFromListEventArgs"/> instance containing the event data.</param>
        protected virtual void OnChooseFromList(ChooseFromListEventArgs e)
        {
            try
            {
                if (this.ChooseFromList != null)
                {
                    this.ChooseFromList(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the  BeforeFormLoad event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeFormLoad(FormEventArgs e)
        {
            try
            {
                if (this.BeforeFormLoad != null)
                {
                    this.BeforeFormLoad(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="BeforeAdd"/> event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeAdd(FormEventArgs e)
        {
            try
            {
                if (this.BeforeAdd != null)
                {
                    this.BeforeAdd(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="BeforeUpdate"/> event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeUpdate(FormEventArgs e)
        {
            try
            {
                if (this.BeforeUpdate != null)
                {
                    this.BeforeUpdate(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="BeforeSave"/> event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeSave(FormEventArgs e)
        {
            try
            {
                if (this.BeforeSave != null)
                {
                    this.BeforeSave(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="OnBeforeDataLoad"/> event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeDataLoad(FormEventArgs e)
        {
            try
            {
                if (this.BeforeDataLoad != null)
                {
                    this.BeforeDataLoad(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="AfterAdd"/> event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterAdd(FormEventArgs e)
        {
            try
            {
                if (this.AfterAdd != null)
                {
                    this.AfterAdd(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="AfterUpdate"/> event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterUpdate(FormEventArgs e)
        {
            try
            {
                if (this.AfterUpdate != null)
                {
                    this.AfterUpdate(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="AfterDataLoad"/> event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterDataLoad(FormEventArgs e)
        {
            try
            {
                if (this.AfterDataLoad != null)
                {
                    this.AfterDataLoad(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="AfterSave"/> event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterSave(FormEventArgs e)
        {
            try
            {
                if (this.AfterSave != null)
                {
                    this.AfterSave(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the AfterFormLoad event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterFormLoad(FormEventArgs e)
        {
            try
            {
                if (this.AfterFormLoad != null)
                {
                    this.AfterFormLoad(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Called when [after arguments set].
        /// </summary>
        protected virtual void OnAfterArgumentsSet()
        {
            try
            {
                if (this.AfterArgumentsSet != null)
                {
                    this.AfterArgumentsSet(this);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the BedoreFormClose event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeFormClose(FormEventArgs e)
        {
            try
            {
                if (this.BeforeFormClose != null)
                {
                    this.BeforeFormClose(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the FormResize event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnFormResize(FormEventArgs e)
        {
            try
            {
                if (this.FormResize != null)
                {
                    this.FormResize(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the BeforeRightClick event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeRightClick(ItemEventArgs e)
        {
            try
            {
                if (this.BeforeRightClick != null)
                {
                    this.BeforeRightClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the AfterRightClick event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterRightClick(ItemEventArgs e)
        {
            try
            {
                if (this.AfterRightClick != null)
                {
                    this.AfterRightClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the BeforeToolbarClick event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeToolbarClick(ItemEventArgs e)
        {
            try
            {
                if (this.BeforeToolbarClick != null)
                {
                    this.BeforeToolbarClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the BeforeMenuClick event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeMenuClick(ItemEventArgs e)
        {
            try
            {
                if (this.BeforeMenuClick != null)
                {
                    this.BeforeMenuClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the AfterToolbarClick event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterToolbarClick(ItemEventArgs e)
        {
            try
            {
                if (this.AfterToolbarClick != null)
                {
                    this.AfterToolbarClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }


        /// <summary>
        /// Raises the AfterMenuClick event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterMenuClick(ItemEventArgs e)
        {
            try
            {
                if (this.AfterMenuClick != null)
                {
                    this.AfterMenuClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the BeforeAddClick event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeAddClick(FormEventArgs e)
        {
            try
            {
                if (this.BeforeAddClick != null)
                {
                    this.BeforeAddClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the <see cref="OnBeforeFindClick"/> event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.UI.Model.EventArguments.FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnBeforeFindClick(FormEventArgs e)
        {
            try
            {
                if (this.BeforeFindClick != null)
                {
                    this.BeforeFindClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the AfterAddClick event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterAddClick(FormEventArgs e)
        {
            try
            {
                if (this.AfterAddClick != null)
                {
                    this.AfterAddClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Raises the AfterFindClick event.
        /// </summary>
        /// <param name="e">The <see cref="FormEventArgs"/> instance containing the event data.</param>
        protected virtual void OnAfterFindClick(FormEventArgs e)
        {
            try
            {
                if (this.AfterFindClick != null)
                {
                    this.AfterFindClick(this, e);
                }
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
            }
        }

        /// <summary>
        /// Opens the query.
        /// </summary>
        /// <param name="sql">The SQL statement.</param>
        /// <param name="queryTitle">The query title.</param>
        /// <returns>True if runs the query successfully.</returns>
        protected bool OpenQuery(string sql, string queryTitle)
        {
            try
            {
                string blankQueryMenu;
                ConfigurationHelper.GlobalConfiguration.Load(this.AddOn.Company, "BlankQueryMenu", out blankQueryMenu);

                if (!string.IsNullOrEmpty(blankQueryMenu))
                {
                    this.AddOn.Application.Menus.Item(blankQueryMenu).Activate();
                    this.AddOn.Application.Forms.ActiveForm.Items.Item("15").Click(BoCellClickType.ct_Regular);
                    this.AddOn.Application.Forms.ActiveForm.Title = queryTitle;
                    ((EditText)this.AddOn.Application.Forms.ActiveForm.Items.Item("3").Specific).Value = sql;
                    this.AddOn.Application.Forms.ActiveForm.Items.Item("12").Click(BoCellClickType.ct_Regular);
                    this.AddOn.Application.Forms.ActiveForm.Items.Item("1").Click(BoCellClickType.ct_Regular);
                }

                return true;
            }
            catch (Exception ex)
            {
                // If an error occurr, log it and fails the method
                this.AddOn.HandleException(ex);
                return false;
            }
        }

        /// <summary>
        /// Opens the query.
        /// </summary>
        /// <param name="sqlHelper">The SQL helper.</param>
        /// <param name="queryTitle">The query title.</param>
        /// <returns>True if runs the query successfully.</returns>
        protected bool OpenQuery(SqlHelper sqlHelper, string queryTitle)
        {
            try
            {
                return this.OpenQuery(sqlHelper.ToString(), queryTitle);
            }
            catch (Exception ex)
            {
                // If an error occurr, log it and fails the method
                this.AddOn.HandleException(ex);
                return false;
            }
        }

        /// <summary>
        /// Exports the specified form.
        /// </summary>
        /// <param name="path">The export path.</param>
        /// <returns>True if the export is successfull.</returns>
        public bool Export(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);                               // Delete the file if it exists.
                }

                // Create the file.
                using (FileStream fs = File.Create(path))
                {
                    string xmlForm = this.SapForm.GetAsXML();

                    // Remove form id and type
                    string pattern = "appformnumber=\"[^\"]+\" ";
                    xmlForm = Regex.Replace(xmlForm, pattern, "appformnumber=\"1000001\" ");
                    pattern = "FormType=\"[^\"]+\" ";
                    xmlForm = Regex.Replace(xmlForm, pattern, "FormType=\"1000001\" ");
                    pattern = "\" uid=\"[^\"]+\" ";
                    xmlForm = Regex.Replace(xmlForm, pattern, "\" uid=\"MCI000001\" ");

                    byte[] info = new UTF8Encoding(true).GetBytes(xmlForm);

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
        /// Determines whether the child is editable
        /// </summary>
        /// <remarks>
        /// The header of the object must be saved <strong>before</strong> any children can be added/modified.
        /// </remarks>
        /// <returns>
        /// 	<c>true</c> if child is editable; otherwise, <c>false</c>.
        /// </returns>
        public bool IsChildEditable()
        {
            try
            {
                switch (this.Mode)
                {
                    case BoFormMode.fm_FIND_MODE:
                        this.AddOn.StatusBar.Message = "This operation is not permitted if the form is in 'Find' mode.";
                        return false;
                    case BoFormMode.fm_ADD_MODE:
                        this.AddOn.StatusBar.Message = "Please add the record before performing this operation.";
                        return false;
                    case BoFormMode.fm_UPDATE_MODE:
                        this.AddOn.StatusBar.Message = "Please update the changes before performing this operation.";
                        return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Applies the changes.
        /// </summary>
        private void ApplyChanges()
        {
            if (this._changesApplied)
            {
                return;
            }

            this._changesApplied = true;

            try
            {
                int currentPane = this.SapForm.PaneLevel;

                var documentXml = new XmlDocument();
                string filePath = string.Format(@"{0}FormsChanges.xml", this.AddOn.Location);

                if (System.IO.File.Exists(filePath))
                {
                    // load the content of the XML File
                    documentXml.Load(filePath);

                    // loop through the form nodes ans remove the one where the FormType does not match
                    var actionNode = documentXml.ChildNodes[1].ChildNodes[0].ChildNodes[0];
                    for (int nodeIndex = actionNode.ChildNodes.Count - 1; nodeIndex >= 0; nodeIndex--)
                    {
                        var formNode = actionNode.ChildNodes[nodeIndex];

                        if (formNode != null && formNode.Attributes != null)
                        {
                            if (formNode.Attributes["FormType"].Value != this.FormTypeId)
                            {
                                actionNode.RemoveChild(formNode);
                            }
                            else
                            {
                                formNode.Attributes["uid"].Value = this.FormId;
                            }
                        }
                    }

                    // load the form to the SBO application in one batch
                    string tmpStr = documentXml.InnerXml.Replace("$AddOnLocation$", this.AddOn.Location);
                    this.AddOn.Application.LoadBatchActions(ref tmpStr);

                    this.SapForm.Freeze(false);
                    this.SapForm.PaneLevel = currentPane;
                    this.SapForm.Freeze(true);

                    this._needsResize = true;
                    this._needsReload = true;
                }
            }
            catch { }
        }

        /// <summary>
        /// Selects the document.
        /// </summary>
        private void SelectDocument()
        {
            try
            {
                if (!string.IsNullOrEmpty(this._findDocumentNumber) && !string.IsNullOrEmpty(this._findDocumentFieldName))
                {
                    this.Mode = BoFormMode.fm_FIND_MODE;
                    this.ControlManager.EditText(this._findDocumentFieldName).Value = this._findDocumentNumber;
                    this._findDocumentNumber = string.Empty;
                    this._findDocumentNumber = string.Empty;
                    this.ControlManager.Button("1").Focus();
                }
            }
            catch
            {
                // Ignore the error in case the field name
                // to find is not an edit box or if
                // some of the controls are not found.
            }
        }

        /// <summary>
        /// Gets the choose from list value.
        /// </summary>
        /// <param name="selectedObjects">The selected objects.</param>
        /// <param name="column">The column.</param>
        /// <returns>The value selected</returns>
        protected string ChooseFromListValue(DataTable selectedObjects, object column)
        {
            try
            {
                string revValue = string.Empty;
                try
                {
                    for (int rowIndex = 0; rowIndex < selectedObjects.Rows.Count; rowIndex++)
                    {
                        if (!string.IsNullOrEmpty(revValue))
                        {
                            revValue += ",";
                        }

                        revValue += selectedObjects.GetValue(column, rowIndex).ToString().Trim();
                    }
                }
                catch
                {
                }

                return revValue;
            }
            catch (Exception ex)
            {
                this.AddOn.HandleException(ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// Chooses from list value.
        /// </summary>
        /// <param name="selectedObjects">The selected objects.</param>
        /// <returns></returns>
        protected string ChooseFromListValue(DataTable selectedObjects)
        {
            return this.ChooseFromListValue(selectedObjects, 0);
        }

        /// <summary>
        /// Gets the primary key.
        /// </summary>
        protected virtual string PrimaryKey
        {
            get
            {
                try
                {
                    string fullKey = this.SapForm.BusinessObject.Key;

                    fullKey = fullKey.Substring(fullKey.IndexOf("DocEntry>") + 9);
                    fullKey = fullKey.Substring(0, fullKey.IndexOf("<"));

                    return fullKey;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// Gets the primary key.
        /// </summary>
        public string DocumentNumber
        {
            get
            {
                try
                {
                    return this.Field(this.TableName, "DocNum").Value;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        #endregion Method
    }
}