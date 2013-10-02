//-----------------------------------------------------------------------
// <copyright file="Button.cs" company="B1C Canada Inc.">
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

    /// <summary>
    /// Instance of the Button class.
    /// </summary>
    /// <author>Frank Alunni</author>
    /// <remarks>The Button adapter represents a Button element in SAP Business One forms. The Button adapter transforms and simplifies the usage of a SAPbouiCOM.Button.</remarks>
    public class Button : Item
    {
        #region Fields

        /// <summary>
        /// The Parent Button
        /// </summary>
        private SAPbouiCOM.Button parentButton;
        #endregion Fields

        #region Constuctors

        /// <summary>
        /// Initializes a new instance of the <see cref="Button"/> class.
        /// </summary>
        public Button()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Button"/> class.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        public Button(IForm form, string uniqueId) : base(form, uniqueId)
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
                    this.parentButton = (SAPbouiCOM.Button)this.BaseItem.Specific;
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

                return this.parentButton.Caption;
            }

            set
            {
                if (this.parentButton == null)
                {
                    return;
                }

                if (this.ParentForm != null)
                {
                    if (!this.ParentForm.Selected)
                    {
                        this.ParentForm.Select();
                    }
                }

                this.parentButton.Caption = value;
            }
        }

        /// <summary>
        /// Gets the base object.
        /// </summary>
        /// <value>The base object.</value>
        public SAPbouiCOM.Button BaseObject
        {
            get { return this.parentButton; }
        }

        #endregion Properties

        #region Indexer
        /// <summary>
        /// Gets or sets the <see cref="System.String"/> with the specified form.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <value>The caption of the button.</value>
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
                    this.parentButton = (SAPbouiCOM.Button)this.BaseItem.Specific;
                }
                else
                {
                    throw new ArgumentNullException(string.Format("Button {0} was not found on form {1}.", uniqueId, form.UniqueID));
                }

                return this.parentButton.Caption;
            }

            set
            {
                if (form == null)
                {
                    return;
                }

                this.ParentForm = form;

                this.BaseItem = form.Items.Item(uniqueId);
                if (this.BaseItem != null)
                {
                    this.parentButton = (SAPbouiCOM.Button)this.BaseItem.Specific;
                }
                else
                {
                    throw new ArgumentNullException(string.Format("Button {0} was not found on form {1}.", uniqueId, form.UniqueID));
                }

                this.parentButton.Caption = value;
            }
        }
        #endregion Indexer

        #region Methods

        /// <summary>
        /// Instances the specified form.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>An instance of the button class</returns>
        public static Button Instance(SAPbouiCOM.Form form, string uniqueId)
        {
            if (form == null)
            {
                return null;
            }

            var button = new Button(form, uniqueId);

            return button.parentButton == null ? null : button;
        }

        /// <summary>
        /// Instances the specified application.
        /// </summary>
        /// <param name="application">The application.</param>
        /// <param name="formUID">The input form UID.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>An instance of the button class</returns>
        public static Button Instance(IApplication application, string formUID, string uniqueId)
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
        /// <param name="value">The input value.</param>
        /// <param name="location">The location.</param>
        /// <param name="size">The button size.</param>
        /// <returns>The instance of the Button class added.</returns>
        public static Button Add(SAPbouiCOM.Form form, string uniqueId, string value, Point location, Size size)
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

            Button button = Instance(form, uniqueId);
            if (button.BaseItem  != null)
            {
                button.Location = location;
                if ((size.Height > 0) && (size.Width > 0))
                {
                    button.Size = size;
                }

                button.Value = value;
            }

            return button;
        }

        /// <summary>
        /// Adds the specified application.
        /// </summary>
        /// <param name="application">The application.</param>
        /// <param name="formUID">The form UID.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <param name="value">The input value.</param>
        /// <param name="location">The location.</param>
        /// <param name="size">The button size.</param>
        /// <returns>The instance of the Button class added.</returns>
        public static Button Add(IApplication application, string formUID, string uniqueId, string value, Point location, Size size)
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
        #endregion Methods
    }
}
