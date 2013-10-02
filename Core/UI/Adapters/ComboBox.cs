//-----------------------------------------------------------------------
// <copyright file="ComboBox.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.UI.Adapters
{
    using System;
    using System.Drawing;
    using Helpers;
    using SAPbouiCOM;
    using B1C.SAP.DI.BusinessAdapters;

    /// <summary>
    /// Instance of the ComboBox class.
    /// </summary>
    /// <author>Frank Alunni</author>
    /// <remarks>The ComboBox adapter represents a Combo box element in SAP Business One forms. It transforms and simplifies the usage of a SAPbouiCOM.ComboBox.</remarks>
    public class ComboBox : Item
    {
        #region Fields
        /// <summary>
        /// The Parent ComboBox
        /// </summary>
        private SAPbouiCOM.ComboBox parentComboBox;
        #endregion Fields

        #region Constuctors
        /// <summary>
        /// Initializes a new instance of the <see cref="ComboBox"/> class.
        /// </summary>
        public ComboBox()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComboBox"/> class.
        /// </summary>
        /// <param name="combo">The combo.</param>
        public ComboBox(SAPbouiCOM.ComboBox combo)
        {
            this.parentComboBox = combo;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComboBox"/> class.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        public ComboBox(IForm form, string uniqueId) : base(form, uniqueId)
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
                    this.parentComboBox = (SAPbouiCOM.ComboBox)this.BaseItem.Specific;
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
        /// Gets or sets the display value.
        /// </summary>
        /// <value>The display value.</value>
        public string DisplayValue
        {
            get
            {
                if (this.ParentForm != null)
                {
                    if (!this.ParentForm.Selected)
                    {
                        this.ParentForm.Select();
                    }
                }

                if (this.parentComboBox == null)
                {
                    return string.Empty;
                }

                if (this.parentComboBox.Selected == null)
                {
                    return string.Empty;
                }

                return this.parentComboBox.Selected.Description;
            }

            set
            {
                if (this.ParentForm != null)
                {
                    if (!this.ParentForm.Selected)
                    {
                        this.ParentForm.Select();
                    }
                }

                if (this.parentComboBox != null)
                {
                    this.parentComboBox.Select(value, BoSearchKey.psk_ByDescription);
                }
            }
        }

        /// <summary>
        /// Gets the base object.
        /// </summary>
        /// <value>The base object.</value>
        public SAPbouiCOM.ComboBox BaseObject
        {
            get { return this.parentComboBox; }
        }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        /// <value>The value.</value>
        public override string Value
        {
            get
            {
                if (this.ParentForm != null)
                {
                    if (!this.ParentForm.Selected)
                    {
                        this.ParentForm.Select();
                    }
                }

                if (this.parentComboBox.Selected == null)
                {
                    return string.Empty;
                }

                return this.parentComboBox.Selected.Value;
            }

            set
            {
                if (this.ParentForm != null)
                {
                    if (!this.ParentForm.Selected)
                    {
                        this.ParentForm.Select();
                    }
                }

                if (this.parentComboBox == null)
                {
                    return;
                }

                if (this.parentComboBox.Selected == null)
                {
                    try
                    {
                        this.parentComboBox.Select(value, BoSearchKey.psk_Index);
                    }
                    catch
                    {
                        this.parentComboBox.Select(value, BoSearchKey.psk_ByValue);
                    }
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
        /// <value>The value of the combobox.</value>
        public string this[SAPbouiCOM.Form form, string uniqueId]
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
                    this.parentComboBox = (SAPbouiCOM.ComboBox)this.BaseItem.Specific;
                }
                else
                {
                    throw new NullReferenceException(string.Format("ComboBox {0} was not found on form {1}.", uniqueId, form.UniqueID));
                }

                if (this.parentComboBox.Selected == null)
                {
                    return string.Empty;
                }

                return this.parentComboBox.Selected.Value;
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
                    this.parentComboBox = (SAPbouiCOM.ComboBox)this.BaseItem.Specific;
                }
                else
                {
                    throw new NullReferenceException(string.Format("ComboBox {0} was not found on form {1}.", uniqueId, form.UniqueID));
                }

                if (this.parentComboBox.Selected != null)
                {
                    this.parentComboBox.Select(value, BoSearchKey.psk_Index);
                }
            }
        }
        #endregion Indexer

        #region Methods
        /// <summary>
        /// Instances the specified form.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>An instance of the combobox class.</returns>
        public static ComboBox Instance(SAPbouiCOM.Form form, string uniqueId)
        {
            if (form == null)
            {
                return null;
            }

            var combo = new ComboBox(form, uniqueId);

            if (combo.parentComboBox == null)
            {
                return null;
            }

            return combo;
        }

        /// <summary>
        /// Instances the specified application.
        /// </summary>
        /// <param name="application">The application.</param>
        /// <param name="formUID">The form UID.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>An instance of the combobox class.</returns>
        public static ComboBox Instance(Application application, string formUID, string uniqueId)
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
        /// <param name="value">The combobox value.</param>
        /// <param name="location">The combobox location.</param>
        /// <param name="size">The combobox size.</param>
        /// <returns>The instance of the combobox added.</returns>
        public static ComboBox Add(SAPbouiCOM.Form form, string uniqueId, string value, Point location, Size size)
        {
            if (form == null)
            {
                return null;
            }

            try
            {
                form.Items.Add(uniqueId, BoFormItemTypes.it_BUTTON);
            }
            catch
            {
            }

            ComboBox control = Instance(form, uniqueId);
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
        /// <param name="value">The combobox value.</param>
        /// <param name="location">The combobox location.</param>
        /// <param name="size">The combobox size.</param>
        /// <returns>The instance of the combobox added.</returns>
        public static ComboBox Add(IApplication application, string formUID, string uniqueId, string value, Point location, Size size)
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
        /// Clears the list.
        /// </summary>
        public void ClearList()
        {
            if (this.parentComboBox != null)
            {
                for (int listIndex = 0; listIndex < this.parentComboBox.ValidValues.Count; listIndex++)
                {
                    this.parentComboBox.ValidValues.Remove(0, BoSearchKey.psk_Index);                    
                }
            }
        }

        /// <summary>
        /// Loads the specified company.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="query">The query.</param>
        /// <param name="displayField">The display field.</param>
        /// <param name="valueField">The value field.</param>
        public void Load(SAPbobsCOM.Company company, string query, string displayField, string valueField)
        {
            this.ClearList();

            using (var rs = new RecordsetAdapter(company, query))
            {
                while (!rs.EoF)
                {
                    string displayValue = rs.FieldValue(displayField).ToString();

                    if (displayValue.Length > 50)
                    {
                        displayValue = displayValue.Substring(0, 47) + "...";
                    }

                    try
                    {
                        this.parentComboBox.ValidValues.Add(rs.FieldValue(valueField).ToString(), displayValue);
                    }
                    catch
                    {
                    }

                    rs.MoveNext();
                }
            }

            if (this.parentComboBox.ValidValues.Count > 0)
            {
                this.parentComboBox.Select(0, BoSearchKey.psk_Index);
            }
        }
        #endregion Methods
    }
}
