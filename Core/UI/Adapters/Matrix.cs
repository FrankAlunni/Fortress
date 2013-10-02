//-----------------------------------------------------------------------
// <copyright file="Matrix.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.UI.Adapters
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using Helpers;
    using B1C.SAP.DI.BusinessAdapters;

    /// <summary>
    /// Instance of the Matrix class.
    /// </summary>
    /// <author>Frank Alunni</author>
    /// <remarks>
    /// The Matrix adapter represents a matrix in SAP Business One application. You use this object to access the matrix columns and cells and to set its properties. It transforms and simplifies the usage of a SAPbouiCOM.Matrix. 
    /// If there are existing rows in the matrix, you will not be able to:
    /// •   Add columns to the matrix 
    /// •   Update the datasource of a column
    /// </remarks>
    public class Matrix : Item
    {
        #region Fields
        /// <summary>
        /// The parent Matrix
        /// </summary>
        private readonly SAPbouiCOM.Matrix parentMatrix;
        #endregion Fields

        #region Constuctors
        /// <summary>
        /// Initializes a new instance of the <see cref="Matrix"/> class.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        public Matrix(SAPbouiCOM.IForm form, string uniqueId) : base(form, uniqueId)
        {
            try
            {
                if (form == null)
                {
                    throw new ArgumentNullException("form");
                }

                this.BaseItem = form.Items.Item(uniqueId);
                if (this.BaseItem != null)
                {
                    this.parentMatrix = (SAPbouiCOM.Matrix)this.BaseItem.Specific;
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
        /// Gets the row count.
        /// </summary>
        /// <value>The row count.</value>
        public virtual int RowCount
        {
            get
            {
                if (this.parentMatrix == null)
                {
                    return 0;
                }

                return this.parentMatrix.RowCount;
            }
        }

        /// <summary>
        /// Gets the columns.
        /// </summary>
        /// <value>The columns.</value>
        public SAPbouiCOM.Columns Columns
        {
            get
            {
                if (this.parentMatrix == null) 
                {
                    return null;
                }

                return this.parentMatrix.Columns;
            }
        }

        /// <summary>
        /// Gets the base object.
        /// </summary>
        /// <value>The base object.</value>
        public SAPbouiCOM.Matrix BaseObject
        {
            get { return this.parentMatrix; }
        }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        /// <value>The string value.</value>
        public override string Value
        {
            get { throw new NotImplementedException(); }

            set { throw new NotImplementedException(); }
        }
        #endregion Properties

        #region Indexers
        /// <summary>
        /// Gets or sets the <see cref="System.String"/> with the specified row.
        /// </summary>
        /// <param name="row">The cell row.</param>
        /// <param name="col">The cell col.</param>
        /// <value>The selected cell.</value>
        public string this[object row, object col]
        {
            get
            {
                if (this.parentMatrix == null) 
                {
                    return string.Empty;
                }

                try
                {
                    var edit = (SAPbouiCOM.EditText)this.Cell(row, col).Specific;
                    return edit.Value;
                }
                catch
                {
                    try
                    {
                        var combo = (SAPbouiCOM.ComboBox)this.Cell(row, col).Specific;
                        return combo.Selected.Value;
                    }
                    catch
                    {
                        try
                        {
                            var check = (SAPbouiCOM.CheckBox)this.Cell(row, col).Specific;
                            return check.Checked ? "Y" : "N";
                        }
                        catch (Exception)
                        {
                            return "Value Not Found";
                        }
                    }
                }
            }

            set
            {
                if (this.parentMatrix != null)
                {
                    var edit = (SAPbouiCOM.EditText)this.Cell(row, col).Specific;
                    edit.Value = value;
                }
            }
        }

        /// <summary>
        /// Gets the Cells with the specified row.
        /// </summary>
        /// <param name="row">The row indentifier.</param>
        /// <value>The selected row.</value>
        public List<SAPbouiCOM.Cell> this[object row]
        {
            get
            {
                if (this.parentMatrix == null) 
                {
                    return null;
                }

                var cells = new List<SAPbouiCOM.Cell>();

                foreach (SAPbouiCOM.Column column in this.parentMatrix.Columns)
                {
                    cells.Add(column.Cells.Item(row));
                }

                return cells;
            }
        }        
        #endregion Indexers

        #region Methods
        #region Public Static
        /// <summary>
        /// Instances the specified form.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <returns>The instance to the Matrix object.</returns>
        public static Matrix Instance(SAPbouiCOM.Form form, string uniqueId)
        {
            return new Matrix(form, uniqueId);
        }

        /// <summary>
        /// Adds the specified form.
        /// </summary>
        /// <param name="form">The input form.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <param name="location">The matrix location.</param>
        /// <param name="size">The matrix size.</param>
        /// <returns>The instance to the Matrix object added.</returns>
        public static Matrix Add(SAPbouiCOM.Form form, string uniqueId, Point location, Size size)
        {
            if (form == null) 
            {
                return null;
            }

            try
            {
                form.Items.Add(uniqueId, SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            }
            catch
            {
            }

            Matrix control = Instance(form, uniqueId);
            if (control != null)
            {
                control.Location = location;
                if ((size.Height > 0) && (size.Width > 0))
                {
                    control.Size = size;
                }
            }

            return control;
        }

        /// <summary>
        /// Adds the specified application.
        /// </summary>
        /// <param name="application">The application.</param>
        /// <param name="formUID">The form UID.</param>
        /// <param name="uniqueId">The unique id.</param>
        /// <param name="location">The matrix location.</param>
        /// <param name="size">The matrix size.</param>
        /// <returns>The instance to the Matrix object added.</returns>
        public static Matrix Add(SAPbouiCOM.IApplication application, string formUID, string uniqueId, Point location, Size size)
        {
            try
            {
                return Add(application.Forms.Item(formUID), uniqueId, location, size);
            }
            catch
            {
                return null;
            }
        }
        #endregion Public Static

        #region Public

        /// <summary>
        /// Fills the width on column.
        /// </summary>
        /// <param name="colToAdjust">
        /// The col to adjust.
        /// </param>
        public void AutoSizeColumn(int colToAdjust)
        {
            if (this.BaseObject == null)
            {
                return;
            }

            int width = 0;
            for (int colIndex = 0; colIndex < this.BaseObject.Columns.Count; colIndex++)
            {
                if (this.BaseObject.Columns.Item(colIndex).Visible && (colIndex != colToAdjust))
                {
                    width += this.BaseObject.Columns.Item(colIndex).Width;
                }
            }

            width += 18; // Scrollbar width

            this.BaseObject.Columns.Item(colToAdjust).Width = this.Width - width;
        }

        /// <summary>
        /// Cells the specified row.
        /// </summary>
        /// <param name="row">The cell row.</param>
        /// <param name="col">The cell col.</param>
        /// <returns>The instance to the Matrix cell.</returns>
        public virtual SAPbouiCOM.CellClass Cell(object row, object col)
        {
            if (this.parentMatrix == null) 
            {
                return null;
            }

            return (SAPbouiCOM.CellClass)this.parentMatrix.Columns.Item(col).Cells.Item(row);
        }

        /// <summary>
        /// Selects the cell.
        /// </summary>
        /// <param name="row">The cell row.</param>
        /// <param name="col">The cell col.</param>
        /// <returns>True if the cell is selected.</returns>
        public virtual bool SelectCell(object row, object col)
        {
            if (this.parentMatrix == null) 
            {
                return false;
            }

            if (this.parentMatrix.RowCount <= 0) 
            {
                return false;
            }

            try
            {
                this.Cell(row, col).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Clears this instance.
        /// </summary>
        public virtual void Clear()
        {
            if (this.parentMatrix != null) 
            {
                this.parentMatrix.Clear();
            }
        }

        /// <summary>
        /// Populates the specified source.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="conditions">The conditions.</param>
        /// <returns>True if the Matrix is populated.</returns>
        public virtual bool Populate(SAPbouiCOM.DBDataSource source, SAPbouiCOM.Conditions conditions)
        {
            if (source == null) 
            {
                return false;
            }

            if (this.parentMatrix == null) 
            {
                return false;
            }

            try
            {
                this.parentMatrix.Clear();               // Ready Matrix to populate data
                source.Query(conditions);           // Querying the DB Data source
                this.parentMatrix.LoadFromDataSource();  // Reload from data source

                return true;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Adds The cell row.
        /// </summary>
        /// <returns>True if the Row is added.</returns>
        public virtual bool AddRow()
        {
            if (this.parentMatrix != null)
            {
                this.parentMatrix.AddRow(1, this.parentMatrix.RowCount);
                this.parentMatrix.SetLineData(this.parentMatrix.RowCount);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Selecteds The cell row.
        /// </summary>
        /// <returns>Returns the row number if the row is selected.</returns>
        public virtual int SelectedRow()
        {
            if (this.parentMatrix == null)
            {
                return -1;
            }

            try
            {
                return this.parentMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// Selects the next row.
        /// </summary>
        /// <returns>Returns the row number if the nect row is selected.</returns>
        public virtual int SelectNextRow()
        {
            try
            {
                int currentRow = this.SelectedRow();
                if (currentRow >= this.parentMatrix.RowCount)
                {
                    return currentRow;
                }

                return this.SelectRow(currentRow + 1);
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// Selects the previous row.
        /// </summary>
        /// <returns>Returns the row number if the previous row is selected.</returns>
        public virtual int SelectPreviousRow()
        {
            try
            {
                int currentRow = this.SelectedRow();
                if (currentRow <= 0)
                {
                    return currentRow;
                }

                return this.SelectRow(currentRow - 1);
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// Selects the last row.
        /// </summary>
        /// <returns>Returns the row number if the last row is selected.</returns>
        public virtual int SelectLastRow()
        {
            try
            {
                return this.SelectRow(this.parentMatrix.RowCount);
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// Selects the first row.
        /// </summary>
        /// <returns>Returns the row number if the first row is selected.</returns>
        public virtual int SelectFirstRow()
        {
            try
            {
                return this.SelectRow(1);
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// Selects The cell row.
        /// </summary>
        /// <param name="row">The cell row.</param>
        /// <returns>Returns the row number if the row is selected.</returns>
        public virtual int SelectRow(int row)
        {
            if (this.parentMatrix == null)
            {
                return -1;
            }

            if (this.parentMatrix.RowCount <= 0)
            {
                return -1;
            }

            try
            {
                this.SelectCell(row, 0);
                return this.SelectedRow();
            }
            catch
            {
                return -1;
            }
        }
        
        /// <summary>
        /// Binds the column.
        /// </summary>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="bound">if set to <c>true</c> [bound].</param>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="alias">The alias.</param>
        /// <returns>True if the column is successfully binded.</returns>
        public virtual bool BindColumn(object columnIndex, bool bound, string tableName, string alias)
        {
            if (this.parentMatrix == null)
            {
                return false;
            }

            if (string.IsNullOrEmpty(alias))
            {
                alias = columnIndex.ToString();
            }

            this.parentMatrix.Columns.Item(columnIndex).DataBind.SetBound(bound, tableName, alias);
            return true;
        }

        /// <summary>
        /// Columns the specified column index.
        /// </summary>
        /// <param name="columnIndex">Index of the column.</param>
        /// <returns>The column object.</returns>
        public virtual SAPbouiCOM.Column Column(object columnIndex)
        {
            if (this.parentMatrix == null)
            {
                return null;
            }

            return this.parentMatrix.Columns.Item(columnIndex);
        }

        /// <summary>
        /// Fills the width on column.
        /// </summary>
        /// <param name="colToAdjust">
        /// The col to adjust.
        /// </param>
        public void FillWidthOnColumn(int colToAdjust)
        {
            if (this.parentMatrix == null)
            {
                return;
            }

            int width = 0;
            for (int colIndex = 0; colIndex < this.parentMatrix.Columns.Count; colIndex++)
            {
                if (this.parentMatrix.Columns.Item(colIndex).Visible && (colIndex != colToAdjust))
                {
                    width += this.parentMatrix.Columns.Item(colIndex).Width;
                }
            }

            width += 18; // Scrollbar width

            this.parentMatrix.Columns.Item(colToAdjust).Width = this.Size.Width - width;
        }

        /// <summary>
        /// Loads the column values.
        /// </summary>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="company">The company.</param>
        /// <param name="query">The query.</param>
        /// <param name="displayField">The display field.</param>
        /// <param name="valueField">The value field.</param>
        public void LoadColumnValues(object columnIndex, SAPbobsCOM.Company company, string query, string displayField, string valueField)
        {
            var combo = this.Columns.Item(columnIndex);
            if (combo == null)
            {
                return;
            }

            while (combo.ValidValues.Count > 0)
            {
                combo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }

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
                        combo.ValidValues.Add(rs.FieldValue(valueField).ToString(), displayValue);
                    }
                    catch
                    {
                    }

                    rs.MoveNext();
                }
            }
        }

        /// <summary>
        /// Deletes the row.
        /// </summary>
        /// <returns></returns>
        public bool DeleteRow()
        {
            if (this.parentMatrix == null)
            {
                return false;
            }

            int selectedRow = this.parentMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
            this.parentMatrix.DeleteRow(selectedRow);

            return true;
        }

        #endregion Public

        #region Protected Static
        #endregion Protected Static

        #region Protected
        #endregion Protected

        #region Private Static
        #endregion Private Static

        #region Private
        #endregion Private

        #endregion Methods
    }
}
