//-----------------------------------------------------------------------
// <copyright file="AdoRecordsetAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters
{
    #region Using Directive(s)
    using Utility.Helpers;
    using System;
    using Utility.Logging;
    using System.Data;
    using System.Data.SqlClient;

    #endregion Using Directive(s)

    /// <summary>
    /// The Recordset Adapter
    /// </summary>
    public class AdoRecordsetAdapter : IDisposable
    {
        /// <summary>
        /// The Sql data reader
        /// </summary>
        private DataSet _dataSet;


        /// <summary>
        /// The Current Row in the DataTable
        /// </summary>
        private int _currentRow = -1;

        /// <summary>
        /// Initializes a new instance of the <see cref="AdoRecordsetAdapter"/> class.
        /// </summary>
        /// <param name="connection">The connection.</param>
        /// <param name="query">The query.</param>
        public AdoRecordsetAdapter(SqlConnection connection, string query)
        {
            try
            {
                var adapter = new SqlDataAdapter(query, connection);
                this._dataSet = new DataSet();

                adapter.Fill(this._dataSet);
                this._currentRow = 0;
            }
            catch (Exception ex)
            {
                ThreadedAppLog.WriteLine("Exception encountered running query: {0}", query);
                throw ex;
            }
        }

        /// <summary>
        /// Gets a value indicating whether end of file.
        /// </summary>
        /// <value><c>true</c> if EOF; otherwise, <c>false</c>.</value>
        public bool EoF
        {
            get { return _currentRow >= this._dataSet.Tables[0].Rows.Count; }
        }

        /// <summary>
        /// Gets a value indicating whether Beginning of File.
        /// </summary>
        /// <value><c>true</c> if BOF; otherwise, <c>false</c>.</value>
        public bool BoF
        {
            get { return _currentRow == 0; }
        }

        /// <summary>
        /// Gets the record count.
        /// </summary>
        /// <value>The record count.</value>
        public int RecordCount
        {
            get
            {
                if (!this.IsLoaded())
                {
                    return -1;
                }

                return this._dataSet.Tables[0].Rows.Count;
            }
        }

        /// <summary>
        /// Executes the specified query.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="query">The query.</param>
        /// <returns>True if executes the specified query.</returns>
        public static bool Execute(SqlConnection connection, string query)
        {
            try
            {
                var command = new SqlCommand(query, connection);
                command.CommandType = CommandType.Text;
                command.ExecuteScalar(); 
                return true;
            }
            catch (Exception ex)
            {
                ThreadedAppLog.WriteLine("Exception encountered running query: {0}", query);
                throw ex;
            }
        }

        /// <summary>
        /// Moves to the first record.
        /// </summary>
        public void MoveFirst()
        {
            if (!this.IsLoaded())
            {
                return;
            }

            this._currentRow = 0;
        }

        /// <summary>
        /// Moves to the last record.
        /// </summary>
        public void MoveLast()
        {
            if (!this.IsLoaded())
            {
                return;
            }

            this._currentRow = this._dataSet.Tables[0].Rows.Count - 1;
        }

        /// <summary>
        /// Moves to the next record.
        /// </summary>
        public void MoveNext()
        {
            this._currentRow++;
        }

        /// <summary>
        /// Moves to the previous record.
        /// </summary>
        public void MovePrevious()
        {
            this._currentRow--;
        }

        /// <summary>
        /// Gets the field value.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <returns>The value of the field with the given index</returns>
        public object FieldValue(object fieldIndex)
        {
            if (!this.IsLoaded())
            {
                return null;
            }

            if (fieldIndex.GetType() == typeof(int))
            {
                return this._dataSet.Tables[0].Rows[this._currentRow][(int)fieldIndex];
            }
            else
            {
                return this._dataSet.Tables[0].Rows[this._currentRow][(string)fieldIndex];
            }

        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
        }

        /// <summary>
        /// Determines whether this instance is loaded.
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if this instance is loaded; otherwise, <c>false</c>.
        /// </returns>
        private bool IsLoaded()
        {
            return (this._dataSet != null && this._dataSet.Tables[0] != null && this._dataSet.Tables[0].Rows.Count > 0 );
        }
    }
}
