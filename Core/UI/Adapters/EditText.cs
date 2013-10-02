//-----------------------------------------------------------------------
// <copyright file="EditText.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.UI.Adapters
{
    using System;
    using System.Drawing;

    /// <summary>
    /// Instance of the EditText class.
    /// </summary>
    /// <author>Frank Alunni</author>
    /// <remarks>The EditText adapter represents a text box item in the SAP Business One forms. It transforms and simplifies the usage of a SAPbouiCOM.EditText.</remarks>
    public class EditText : Item
    {
        #region Fields
        
        /// <summary>
        /// The Parent EditText
        /// </summary>
        private SAPbouiCOM.EditText parentEditText;
        #endregion Fields

        #region Constuctors

        /// <summary>
        /// Initializes a new instance of the <see cref="EditText"/> class.
        /// </summary>
        public EditText()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EditText"/> class.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        public EditText(SAPbouiCOM.IForm form, string uniqueId) : base(form, uniqueId)
        {
            try
            {
                if (form == null)
                {
                    return;
                }

                this.BaseItem = form.Items.Item(uniqueId);
                if (this.BaseItem != null)
                {
                    this.parentEditText = (SAPbouiCOM.EditText)this.BaseItem.Specific;
                }
            }
            catch
            {
                return;
            }
        }
        #endregion Constuctors

        #region Properties

        /// <summary>
        /// Gets the base object.
        /// </summary>
        /// <value>The base object.</value>
        public SAPbouiCOM.EditText BaseObject
        {
            get { return this.parentEditText; }
        }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        /// <value>The value.</value>
        public override string Value
        {
            get
            {
                if (this.ParentForm == null)
                {
                    return string.Empty;
                }

                if (this.parentEditText == null)
                {
                    return string.Empty;
                }

                // Try to get the value from the databind
                try
                {
                    string table = this.parentEditText.DataBind.TableName;
                    string field = this.parentEditText.DataBind.Alias;

                    if (string.IsNullOrEmpty(table))
                    {
                        // User Data source
                        var sourceUser = this.ParentForm.DataSources.UserDataSources.Item(field);
                        return sourceUser.Value;
                    }

                    // DB Data Source
                    var sourceDatabase = this.ParentForm.DataSources.DBDataSources.Item(table);
                    return sourceDatabase.GetValue(field, sourceDatabase.Offset).Trim(' ');
                }
                catch 
                {
                }

                try
                {
                    if (!this.ParentForm.Selected)
                    {
                        this.ParentForm.Select();
                    }

                    this.parentEditText.Active = true;
                }
                catch
                {
                }

                return this.parentEditText.Value.Trim(' ');
            }

            set
            {
                if (this.parentEditText == null)
                {
                    return;
                }

                try
                {
                    if (this.ParentForm != null)
                    {
                        if (!this.ParentForm.Selected)
                        {
                            this.ParentForm.Select();
                        }
                    }

                    this.parentEditText.Active = true;
                }
                catch
                {
                }

                try
                {
                    this.parentEditText.Value = value;
                }
                catch
                {
                }
            }
        }

        #endregion Properties

        #region Indexer
        /// <summary>
        /// Gets or sets the <see cref="System.String"/> with the specified form.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <value>The value of the edittext</value>
        public string this[SAPbouiCOM.IForm form, string uniqueId]
        {
            get
            {
                if (form == null)
                {
                    return null;
                }

                this.ParentForm = form;
                this.UniqueId = uniqueId;

                this.BaseItem = form.Items.Item(uniqueId);
                if (this.BaseItem != null)
                {
                    this.parentEditText = (SAPbouiCOM.EditText)this.BaseItem.Specific;
                }
                else
                {
                    throw new NullReferenceException(string.Format("EditText {0} was not found on form {1}.", uniqueId, form.UniqueID));
                }

                return this.parentEditText.Value;
            }

            set
            {
                if (form == null)
                {
                    return;
                }

                this.ParentForm = form;
                this.UniqueId = uniqueId;

                this.BaseItem = form.Items.Item(uniqueId);
                if (this.BaseItem != null)
                {
                    this.parentEditText = (SAPbouiCOM.EditText)this.BaseItem.Specific;
                }
                else
                {
                    throw new NullReferenceException(string.Format("EditText {0} was not found on form {1}.", uniqueId, form.UniqueID));
                }

                this.parentEditText.Value = value;
            }
        }
        #endregion Indexer

        #region Methods
        /// <summary>
        /// Instances the specified EditText.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>The Instance to an EdiText.</returns>
        public static EditText Instance(SAPbouiCOM.Form form, string uniqueId)
        {
            return new EditText(form, uniqueId);
        }

        /// <summary>
        /// Instances the specified EditText.
        /// </summary>
        /// <param name="application">The application.</param>
        /// <param name="formUID">The form UID.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>The Instance to an EdiText.</returns>
        public static EditText Instance(SAPbouiCOM.Application application, string formUID, string uniqueId)
        {
            try
            {
                return Instance(application.Forms.Item(formUID), uniqueId);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Adds the specified form.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <param name="value">The edittext value.</param>
        /// <param name="location">The edittext location.</param>
        /// <param name="size">The edittext size.</param>
        /// <returns>The EditText added.</returns>
        public static EditText Add(SAPbouiCOM.Form form, string uniqueId, string value, Point location, Size size)
        {
            if (form == null)
            {
                return null;
            }

            try
            {
                form.Items.Add(uniqueId, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            }
            catch 
            { 
            }

            EditText control = Instance(form, uniqueId);
            if (control != null)
            {
                control.Location = location;
                if ((size.Height > 0) && (size.Width > 0))
                {
                    control.Size = size;
                }

                control.Value = value;
            }

            return control;
        }

        /// <summary>
        /// Adds the specified application.
        /// </summary>
        /// <param name="application">The application.</param>
        /// <param name="formUID">The form UID.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <param name="value">The edittext value.</param>
        /// <param name="location">The edittext location.</param>
        /// <param name="size">The edittext size.</param>
        /// <returns>The EditText added.</returns>
        public static EditText Add(SAPbouiCOM.Application application, string formUID, string uniqueId, string value, Point location, Size size)
        {
            try
            {
                return Add(application.Forms.Item(formUID), uniqueId, value, location, size);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Selecteds the text.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="app">The application object .</param>
        /// <returns>The text selected.</returns>
        public string SelectedText(SAPbouiCOM.Form form, SAPbouiCOM.Application app)
        {
            try
            {
                if (this.BaseObject == null)
                {
                    return string.Empty;
                }

                if (app == null)
                {
                    return string.Empty;
                }

                if (form == null)
                {
                    return string.Empty;
                }

                this.ParentForm = form;

                try
                {
                    app.ActivateMenuItem("772");    // Copy into clipboard
                }
                catch
                {
                    return string.Empty;  // There is nothing to copy
                }

                string clipboardId = string.Format("x{0}xxxxxxxx", BaseItem.UniqueID).Substring(0, 8);
                
                SAPbouiCOM.Item clipboardItem;
                try
                {
                    clipboardItem = form.Items.Add(clipboardId, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                }
                catch
                {
                    clipboardItem = form.Items.Item(clipboardId);
                }

                clipboardItem.Width = form.Width;
                clipboardItem.Height = 1;
                var clipboardEdit = (SAPbouiCOM.EditText)clipboardItem.Specific;

                // Set focus
                clipboardItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                clipboardEdit.Value = string.Empty; // Cleat the current content
                app.ActivateMenuItem("773");    // Paste from clipboard
                string selectedValue = clipboardEdit.Value;
                clipboardEdit.Value = string.Empty;
                
                // Reset focus back
                clipboardItem.Width = 0;
                clipboardItem.Height = 0;
                BaseItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                return selectedValue;
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion Methods
    }
}
