//-----------------------------------------------------------------------
// <copyright file="Item.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.UI.Adapters
{
    using System.Drawing;

    /// <summary>
    /// Instance of the Item class.
    /// </summary>
    /// <author>Frank Alunni</author>
    /// <remarks>The base adapter component exposes properties and methods common to all the control adapters developed using the core Add-On. The form adapter is an exception due to the different nature of the SAPbouiCOM.Form object encapsulated in it.</remarks>
    public class Item
    {
        #region Fields

        /// <summary>
        /// The Parent SAP UI Item for the object
        /// </summary>
        private SAPbouiCOM.Item parentItem;
        #endregion Fields

        /// <summary>
        /// Initializes a new instance of the <see cref="Item"/> class.
        /// </summary>
        public Item()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Item"/> class.
        /// </summary>
        /// <param name="form">The current form.</param>
        /// <param name="uniqueId">The unique id.</param>
        public Item(SAPbouiCOM.IForm form, string uniqueId)
        {
            try
            {
                if (form == null)
                {
                    return;
                }

                this.ParentForm = form;
                this.UniqueId = uniqueId;
                this.parentItem = this.ParentForm.Items.Item(this.UniqueId);
            }
            catch
            {
                return;
            }
        }

        #region Properties
        /// <summary>
        /// Gets or sets the parent form.
        /// </summary>
        /// <value>The parent form.</value>
        public SAPbouiCOM.IForm ParentForm { get; set; }

        /// <summary>
        /// Gets or sets the unique id.
        /// </summary>
        /// <value>The unique id.</value>
        public string UniqueId { get; set; }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        /// <value>The value.</value>
        public virtual string Value { get; set; }

        /// <summary>
        /// Gets or sets the item.
        /// </summary>
        /// <value>The item object.</value>
        public SAPbouiCOM.Item BaseItem
        {
            get { return this.parentItem; }

            set { this.parentItem = value; }
        }

        /// <summary>
        /// Gets or sets the location.
        /// </summary>
        /// <value>The location.</value>
        public virtual Point Location
        {
            get 
            {
                return new Point(this.parentItem.Left, this.parentItem.Top); 
            }

            set
            {
                this.parentItem.Top = value.Y;
                this.parentItem.Left = value.X;
            }
        }

        /// <summary>
        /// Gets or sets the current panel.
        /// </summary>
        /// <value>The current panel.</value>
        public virtual int CurrentPanel
        {
            get
            {
                return this.parentItem.FromPane;
            }

            set
            {
                this.parentItem.FromPane = value;
                this.parentItem.ToPane = value;
            }
        }

        /// <summary>
        /// Gets or sets the size.
        /// </summary>
        /// <value>The size of the item.</value>
        public virtual Size Size
        {
            get
            {
                return new Size(this.parentItem.Width, this.parentItem.Height);
            }

            set
            {
                this.parentItem.Width = value.Width;
                this.parentItem.Height = value.Height;
            }
        }

        /// <summary>
        /// Gets or sets the width.
        /// </summary>
        /// <value>The control width.</value>
        public virtual int Width
        {
             get
             {
                 return this.parentItem.Width;                 
             }

            set
            {
                this.parentItem.Width = value;
            }
        }

        /// <summary>
        /// Gets or sets the height.
        /// </summary>
        /// <value>The control height.</value>
        public virtual int Height
        {
            get
            {
                return this.parentItem.Height;
            }

            set
            {
                this.parentItem.Height = value;
            }
        }

        /// <summary>
        /// Gets or sets the left coordinate.
        /// </summary>
        /// <value>The control left coordinate.</value>
        public virtual int Left
        {
            get
            {
                return this.parentItem.Left;
            }

            set
            {
                this.parentItem.Left = value;
            }
        }


        /// <summary>
        /// Gets or sets the top coordinate.
        /// </summary>
        /// <value>The control top coordinate.</value>
        public virtual int Top
        {
            get
            {
                return this.parentItem.Top;
            }

            set
            {
                this.parentItem.Top = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="Item"/> is enabled.
        /// </summary>
        /// <value><c>true</c> if enabled; otherwise, <c>false</c>.</value>
        public bool Enabled
        {
            get { return this.parentItem.Enabled; }

            set { this.parentItem.Enabled = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="Item"/> is visible.
        /// </summary>
        /// <value><c>true</c> if visible; otherwise, <c>false</c>.</value>
        public bool Visible
        {
            get { return this.parentItem.Visible; }

            set { this.parentItem.Visible = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether [affects form mode].
        /// </summary>
        /// <value><c>true</c> if [affects form mode]; otherwise, <c>false</c>.</value>
        public bool AffectsFormMode
        {
            get { return this.parentItem.AffectsFormMode; }

            set { this.parentItem.AffectsFormMode = value; }
        }
        #endregion Properties

        #region Methods

        /// <summary>
        /// Moves the specified top.
        /// </summary>
        /// <param name="top">The top coordinate.</param>
        /// <param name="left">The left coordinate.</param>
        public virtual void Move(int top, int left)
        {
            this.BaseItem.Top = top;
            this.BaseItem.Left = left;
        }

        /// <summary>
        /// Moves the on right of.
        /// </summary>
        /// <param name="control">The control.</param>
        public virtual void MoveOnRightOf(Item control)
        {
            this.MoveOnRightOf(control, 5);
        }

        /// <summary>
        /// Moves the on right of.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="pixels">The pixels.</param>
        public virtual void MoveOnRightOf(Item control, int pixels)
        {
            this.BaseItem.Top = control.BaseItem.Top;
            this.BaseItem.Left = control.BaseItem.Left + control.BaseItem.Width + pixels;
        }

        /// <summary>
        /// Moves the on left of.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="pixels">The pixels.</param>
        public virtual void MoveOnLeftOf(Item control, int pixels)
        {
            this.BaseItem.Top = control.BaseItem.Top;
            this.BaseItem.Left = control.BaseItem.Left - (control.BaseItem.Width + pixels);
        }

        /// <summary>
        /// Moves the on left of.
        /// </summary>
        /// <param name="control">The control.</param>
        public virtual void MoveOnLeftOf(Item control)
        {
            this.MoveOnLeftOf(control, 5);
        }

        /// <summary>
        /// Moves the above.
        /// </summary>
        /// <param name="control">The control.</param>
        public virtual void MoveAbove(Item control)
        {
            this.MoveAbove(control, 5);
        }

        /// <summary>
        /// Moves the above.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="pixels">The pixels.</param>
        public virtual void MoveAbove(Item control, int pixels)
        {
            this.BaseItem.Top = control.BaseItem.Top - (control.BaseItem.Height + pixels);
            this.BaseItem.Left = control.BaseItem.Left;
        }

        /// <summary>
        /// Moves the inside.
        /// </summary>
        /// <param name="control">The control.</param>
        public virtual void MoveInside(Item control)
        {
            this.MoveInside(control, 5);
        }

        /// <summary>
        /// Moves the inside.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="pixels">The pixels.</param>
        public virtual void MoveInside(Item control, int pixels)
        {
            this.BaseItem.Top = control.BaseItem.Top + pixels;
            this.BaseItem.Left = control.BaseItem.Left + pixels;
            this.BaseItem.Height = control.BaseItem.Height - (pixels * 2);
            this.BaseItem.Width = control.BaseItem.Width - (pixels * 2);
        }

        /// <summary>
        /// Moves the Outside.
        /// </summary>
        /// <param name="control">The control.</param>
        public virtual void MoveOutside(Item control)
        {
            this.MoveOutside(control, 5);
        }

        /// <summary>
        /// Moves the Outside.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="pixels">The pixels.</param>
        public virtual void MoveOutside(Item control, int pixels)
        {
            this.BaseItem.Top = control.BaseItem.Top - pixels;
            this.BaseItem.Left = control.BaseItem.Left - pixels;
            this.BaseItem.Height = control.BaseItem.Height + (pixels * 2);
            this.BaseItem.Width = control.BaseItem.Width + (pixels * 2);
        }

        /// <summary>
        /// Sizes as control.
        /// </summary>
        /// <param name="control">The control.</param>
        public virtual void SizeAs(Item control)
        {
            this.BaseItem.Width = control.BaseItem.Width;
            this.BaseItem.Height = control.BaseItem.Height;
        }

        /// <summary>
        /// Moves the below.
        /// </summary>
        /// <param name="control">The control.</param>
        public virtual void MoveBelow(Item control)
        {
            this.MoveBelow(control, 5);
        }

        /// <summary>
        /// Moves the below.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="pixels">The pixels.</param>
        public virtual void MoveBelow(Item control, int pixels)
        {
            this.BaseItem.Top = control.BaseItem.Top + (control.BaseItem.Height + pixels);
            this.BaseItem.Left = control.BaseItem.Left;
        }

        /// <summary>
        /// Moves the on top of.
        /// </summary>
        /// <param name="control">The control.</param>
        public virtual void MoveOnTopOf(Item control)
        {
            this.SizeAs(control);
            this.Location = control.Location;
            control.Visible = false;
        }

        /// <summary>
        /// Removes the specified application.
        /// </summary>
        /// <param name="application">The application.</param>
        /// <param name="formUID">The form UID.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>True of the method removes successfully.</returns>
        public bool Remove(SAPbouiCOM.Application application, string formUID, string uniqueId)
        {
            try
            {
                return this.Remove(application.Forms.Item(formUID), uniqueId);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Removes the specified form.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>True of the method removes successfully.</returns>
        public bool Remove(SAPbouiCOM.Form form, string uniqueId)
        {
            try
            {
                form.Items.Item(uniqueId).Visible = false;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Removes this instance.
        /// </summary>
        /// <returns>True of the method removes successfully.</returns>
        public bool Remove()
        {
            try
            {
                this.parentItem.Visible = false;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Focuses this instance.
        /// </summary>
        public void Focus()
        {
            if (this.parentItem == null)
            {
                return;
            }

            this.parentItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        }
        #endregion Methods
    }
}
