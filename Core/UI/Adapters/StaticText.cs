//-----------------------------------------------------------------------
// <copyright file="StaticText.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.UI.Adapters
{
    using System;
    using System.Drawing;
    using Helpers;
    
    /// <summary>
    /// Adapter for a SAPbouiCOM.StaticText
    /// </summary>
    /// <remarks>The StaticText adapter represents a label in SAP Business One application. The StaticText adapter transforms and simplifies the usage of a SAPbouiCOM.StaticText</remarks>
    public class StaticText : Item
    {
        #region Fields
        /// <summary>
        /// The Parent StaticText
        /// </summary>
        private SAPbouiCOM.StaticText parentStaticText;
        #endregion Fields

        #region Constuctors
        /// <summary>
        /// Initializes a new instance of the <see cref="StaticText"/> class.
        /// </summary>
        public StaticText()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="StaticText"/> class.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        public StaticText(SAPbouiCOM.IForm form, string uniqueId) : base(form, uniqueId)
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
                    this.parentStaticText = (SAPbouiCOM.StaticText)this.BaseItem.Specific;
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

                return this.parentStaticText.Caption;
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

                if (this.parentStaticText == null)
                {
                    return;
                }

                this.parentStaticText.Caption = value;
            }
        }
        #endregion Properties

        #region Indexer
        /// <summary>
        /// Gets or sets the <see cref="System.String"/> with the specified form.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <value>The text value of the statictext</value>
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
                    this.parentStaticText = (SAPbouiCOM.StaticText)this.BaseItem.Specific;
                }
                else
                {
                    throw new NullReferenceException(string.Format("StaticText {0} was not found on form {1}.", uniqueId, form.UniqueID));
                }

                return this.parentStaticText.Caption;
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
                    this.parentStaticText = (SAPbouiCOM.StaticText)this.BaseItem.Specific;
                }
                else
                {
                    throw new ArgumentNullException(string.Format("StaticText {0} was not found on form {1}.", uniqueId, form.UniqueID));
                }

                this.parentStaticText.Caption = value;
            }
        }
        #endregion Indexer

        #region Methods
        #region Public Static
        /// <summary>
        /// Instances the specified form.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>An instance of the statictext.</returns>
        public static StaticText Instance(SAPbouiCOM.Form form, string uniqueId)
        {
            return new StaticText(form, uniqueId);
        }

        /// <summary>
        /// Instances the specified application.
        /// </summary>
        /// <param name="application">The application.</param>
        /// <param name="formUID">The form UID.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>An instance of the statictext.</returns>
        public static StaticText Instance(SAPbouiCOM.Application application, string formUID, string uniqueId)
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
        /// <param name="value">The statictext value.</param>
        /// <param name="location">The statictext location.</param>
        /// <param name="size">The statictext size.</param>
        /// <returns>An instance of the statictext added.</returns>
        public static StaticText Add(SAPbouiCOM.Form form, string uniqueId, string value, Point location, Size size)
        {
            if (form == null)
            {
                return null;
            }

            try
            {
                form.Items.Add(uniqueId, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            }
            catch
            { 
            }

            StaticText control = Instance(form, uniqueId);

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
        /// <param name="value">The statictext value.</param>
        /// <param name="location">The statictext location.</param>
        /// <param name="size">The statictext size.</param>
        /// <returns>An instance of the statictext added.</returns>
        public static StaticText Add(SAPbouiCOM.Application application, string formUID, string uniqueId, string value, Point location, Size size)
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

        #endregion Public Static

        #endregion Methods
    }
}
