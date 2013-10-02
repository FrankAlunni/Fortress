using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using B1C.SAP.DI.Model.Exceptions;
using System.Runtime.InteropServices;

namespace B1C.SAP.DI.BusinessAdapters
{
    public class TableRegistration : SapObjectAdapter
    {
        #region Properties
        /// <summary>
        /// Gets or sets the name of the table.
        /// </summary>
        /// <value>The name of the table.</value>
        public string TableName { get; set; }
        /// <summary>
        /// Gets or sets the name of the object.
        /// </summary>
        /// <value>The name of the object.</value>
        public string ObjectName { get; set; }

        /// <summary>
        /// Gets or sets the object code.
        /// </summary>
        /// <value>The object code.</value>
        public string ObjectCode { get; set; }

        /// <summary>
        /// Gets or sets the type of the object.
        /// </summary>
        /// <value>The type of the object.</value>
        public BoUDOObjType ObjectType { get; set; }

        /// <summary>
        /// Gets or sets the child tables CSV.
        /// </summary>
        /// <value>The child tables CSV.</value>
        public List<string> ChildTables { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance can delete.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance can delete; otherwise, <c>false</c>.
        /// </value>
        public bool CanDelete { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance can cancel.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance can cancel; otherwise, <c>false</c>.
        /// </value>
        public bool CanCancel { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance can close.
        /// </summary>
        /// <value><c>true</c> if this instance can close; otherwise, <c>false</c>.</value>
        public bool CanClose { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance can create default form.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance can create default form; otherwise, <c>false</c>.
        /// </value>
        public bool CanCreateDefaultForm { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance can find.
        /// </summary>
        /// <value><c>true</c> if this instance can find; otherwise, <c>false</c>.</value>
        public bool CanFind { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance can log.
        /// </summary>
        /// <value><c>true</c> if this instance can log; otherwise, <c>false</c>.</value>
        public bool CanLog { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance can year transfer.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance can year transfer; otherwise, <c>false</c>.
        /// </value>
        public bool CanYearTransfer { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether [manage series].
        /// </summary>
        /// <value><c>true</c> if [manage series]; otherwise, <c>false</c>.</value>
        public bool ManageSeries { get; set; }

        /// <summary>
        /// Gets or sets the find columns CSV.
        /// </summary>
        /// <value>The find columns CSV.</value>
        public List<string> FindColumns { get; set; }

        /// <summary>
        /// Gets or sets the form columns CSV.
        /// </summary>
        /// <value>The form columns CSV.</value>
        public List<string> FormColumns { get; set; }
        #endregion

        #region Constructors
        public TableRegistration(Company company, string tableName)
            : base(company)
        {
            this.TableName = tableName;
            this.ChildTables = new List<string>();
            this.FindColumns = new List<string>();
            this.FormColumns = new List<string>();
        }
        #endregion

        #region Methods
        /// <summary>
        /// Adds this instance.
        /// </summary>
        public void Add()
        {
            // may not be a user table
            SAPbobsCOM.Recordset rs = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                string query = string.Format("SELECT * FROM OUDO WHERE Code = '{0}'", this.ObjectCode);
                rs.DoQuery(query);
                if (!rs.EoF)
                {
                    return;
                }
            }
            catch
            {
                // Ingore error table does not exists
            }
            finally
            {
                Marshal.ReleaseComObject(rs);
                rs = null;
                GC.Collect();
            }

            var reg = (UserObjectsMD)this.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

            try
            {
                reg.CanLog = this.CanLog ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;
                reg.CanCancel = this.CanCancel ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;
                reg.CanClose = this.CanClose ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;
                reg.CanDelete = this.CanDelete ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;
                reg.CanYearTransfer = this.CanYearTransfer ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;
                reg.ManageSeries = this.ManageSeries ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;
                reg.Code = this.ObjectCode;
                reg.Name = this.ObjectName;
                reg.ObjectType = this.ObjectType;
                reg.TableName = this.TableName;

                if (this.CanLog)
                {
                    reg.LogTableName = "A" + this.TableName;
                }

                var childIndex = 0;
                foreach (var childTable in this.ChildTables)
                {
                    if (childIndex >= reg.ChildTables.Count)
                    {
                        reg.ChildTables.Add();
                    }

                    reg.ChildTables.SetCurrentLine(childIndex);
                    reg.ChildTables.TableName = childTable;

                    childIndex++;
                }

                // Add the find columns
                if (this.CanFind)
                {
                    var columnIndex = 0;
                    foreach (var columnName in this.FindColumns)
                    {
                        if (columnIndex >= reg.FindColumns.Count)
                        {
                            reg.FindColumns.Add();
                        }

                        reg.FindColumns.SetCurrentLine(columnIndex);
                        reg.FindColumns.ColumnAlias = columnName;

                        columnIndex++;
                    }

                    reg.CanFind = BoYesNoEnum.tYES;
                }
                else
                {
                    reg.CanFind = BoYesNoEnum.tNO;
                }


                // Add the form columns
                if (this.CanCreateDefaultForm)
                {
                    var columnIndex = 0;
                    foreach (var columnName in this.FormColumns)
                    {
                        if (columnIndex >= reg.FormColumns.Count)
                        {
                            reg.FormColumns.Add();
                        }

                        reg.FormColumns.SetCurrentLine(columnIndex);
                        reg.FormColumns.FormColumnAlias = columnName;

                        columnIndex++;
                    }

                    reg.CanCreateDefaultForm = BoYesNoEnum.tYES;
                }
                else
                {
                    reg.CanCreateDefaultForm = BoYesNoEnum.tNO;
                }

                if (reg.Add() != 0)
                {
                    throw new SapException(this.Company);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(reg);
                GC.Collect();
            }
        }

        protected override void Release()
        {
        }
        #endregion

    }
}
