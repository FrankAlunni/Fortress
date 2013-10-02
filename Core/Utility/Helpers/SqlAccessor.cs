// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SqlAccessor.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   The DataHelpers class provides an easy-to-use interface to running
//   SQL, and executing stored procedures. The class also supports transactions using the SqlTransaction
//   object that is a part of ADO.NET.
//   <newpara></newpara>As a developer using this class, you do not need to
//   worry about opening and closing connections, for that is handled within the methods
//   behind the scenes. These features will provide automatic database connection pooling, which is an
//   important part of a scalable application.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.Utility.Helpers
{
    using System;
    using System.Collections;
    using System.Configuration;
    using System.Data;
    using System.Data.SqlClient;

    /// <summary>
    /// The DataHelpers class provides an easy-to-use interface to running
    /// SQL, and executing stored procedures. The class also supports transactions using the SqlTransaction
    /// object that is a part of ADO.NET.
    /// <newpara></newpara>As a developer using this class, you do not need to
    /// worry about opening and closing connections, for that is handled within the methods
    /// behind the scenes. These features will provide automatic database connection pooling, which is an
    /// important part of a scalable application.
    /// </summary>
    public sealed class SqlAccessor
    {
        #region Fields

        /// <summary>
        /// The isolation level that all operations for the 
        /// instance should run in.
        /// </summary>
        private readonly IsolationLevel isoLevel = IsolationLevel.ReadUncommitted;

        /// <summary>
        /// A global connection to be used in methods that require a
        /// connection, if the method was not supplied a connection to
        /// use.
        /// </summary>
        private SqlConnection connection;

        /// <summary>
        /// The database connection string 
        /// </summary>
        private string connectionString = string.Empty;

        /// <summary>
        /// A global transaction to be used with 'connection'
        /// in methods that can use a transaction, if the method was
        /// not supplied a transaction to use.
        /// </summary>
        private SqlTransaction transaction;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlAccessor"/> class. 
        /// Constructor that allows for the setting of the connectionStringKey property
        /// of this object instance.
        /// </summary>
        /// <param name="connectionStringKey">
        /// The key to be used when obtaining
        /// the connection string from the configuration files. Ex: VaultConnectionString.
        /// </param>
        public SqlAccessor(string connectionStringKey)
        {
            this.CommandTiemout = 60;
            this.connectionString = string.Empty;
            this.ConnectionStringKey = connectionStringKey; 
            this.isoLevel = IsolationLevel.ReadUncommitted;
        }

        #endregion 

        #region Enums

        #region Public

        #region DataTypes enum

        /// <summary>
        /// The DataTypes enumeration provides a list of valid SQL data types for a parameter 
        /// to be used in a data function.  If you are trying to create a parameter that contains
        /// a large integer value, the IntDT value should be used.
        /// </summary>
        public enum DataTypes
        {
            /// <summary>
            /// Encapsulates System.Data.SqlDbType.BigInt
            /// </summary>
            BigIntDT = SqlDbType.BigInt, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Binary
            /// </summary>
            BinaryDT = SqlDbType.Binary, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Bit
            /// </summary>
            BitDT = SqlDbType.Bit, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Bit
            /// </summary>
            BoolDT = SqlDbType.Bit, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Char
            /// </summary>
            CharDT = SqlDbType.NChar, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.DateTime
            /// </summary>
            DateTimeDT = SqlDbType.DateTime, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Decimal
            /// </summary>
            DecimalDT = SqlDbType.Decimal, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Float
            /// </summary>
            FloatDT = SqlDbType.Float, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Image
            /// </summary>
            ImageDT = SqlDbType.Image, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Int
            /// </summary>
            IntDT = SqlDbType.Int, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Money
            /// </summary>
            MoneyDT = SqlDbType.Money, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.NChar
            /// </summary>
            NCharDT = SqlDbType.NChar, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.NText
            /// </summary>
            NTextDT = SqlDbType.NText, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.NVarChar
            /// </summary>
            NVarCharDT = SqlDbType.NVarChar, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Real
            /// </summary>
            RealDT = SqlDbType.Real, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.SmallDateTime
            /// </summary>
            SmallDateTimeDT = SqlDbType.SmallDateTime, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.SmallInt
            /// </summary>
            SmallIntDT = SqlDbType.SmallInt, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.SmallMoney
            /// </summary>
            SmallMoneyDT = SqlDbType.SmallMoney, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Text
            /// </summary>
            TextDT = SqlDbType.NText, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Timestamp
            /// </summary>
            TimeStampDT = SqlDbType.Timestamp, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.TinyInt
            /// </summary>
            TinyIntDT = SqlDbType.TinyInt, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.UniqueIdentifier
            /// </summary>
            UniqueIdentifierDT = SqlDbType.UniqueIdentifier, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.VarBinary
            /// </summary>
            VarBinaryDT = SqlDbType.VarBinary, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.VarChar
            /// </summary>
            VarCharDT = SqlDbType.VarChar, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Variant
            /// </summary>
            VariantDT = SqlDbType.Variant, 

            /// <summary>
            /// Encapsulates System.Data.SqlDbType.Xml
            /// </summary>
            XmlDT = SqlDbType.Xml
        }

        #endregion

        #region ParamDirections enum

        /// <summary>
        /// The ParamDirections enumeration provides a list of direction options for a parameter that 
        /// is to be placed into a stored procedure.  For example, a parameter that is meant to
        /// be a return value should be set to ReturnValuePD
        /// </summary>
        public enum ParamDirections
        {
            /// <summary>
            /// Encapsulates System.Data.ParameterDirection.Input
            /// </summary>
            InputPD = ParameterDirection.Input, 

            /// <summary>
            /// Encapsulates System.Data.ParameterDirection.InputOutput
            /// </summary>
            InputOutputPD = ParameterDirection.InputOutput, 

            /// <summary>
            /// Encapsulates System.Data.ParameterDirection.OutputPD
            /// </summary>
            OutputPD = ParameterDirection.Output, 

            /// <summary>
            /// Encapsulates System.Data.ParameterDirection.ReturnValue
            /// </summary>
            ReturnValuePD = ParameterDirection.ReturnValue
        }

        #endregion

        #endregion Public

        #endregion 

        #region Properties

        #region Public

        /// <summary>
        /// Gets or sets the connection string key.
        /// The key used in the configuration files to look up the 
        /// connection string for the database operation.
        /// </summary>
        /// <value>The connection string key.</value>
        public string ConnectionStringKey { get; set; }

        /// <summary>
        /// Gets or sets the query timeout time on an SqlCommand
        /// </summary>
        public int CommandTiemout { get; set; }

        #endregion Public

        #endregion Properties

        #region Methods

        /// <summary>
        /// Gets the string value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the string value</returns>
        public static string GetStringValue(DataRowView row, string fieldName)
        {
            return GetStringValue(row, fieldName, string.Empty);
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the string value</returns>
        public static string GetStringValue(SqlDataReader reader, string fieldName)
        {
            return GetStringValue(reader, fieldName, string.Empty);
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns>Returns the string value</returns>
        public static string GetStringValue(DataRowView row, string fieldName, string defaultValue)
        {
            if (row[fieldName] != DBNull.Value)
            {
                return (string)row[fieldName];
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns>Returns the string value</returns>
        public static string GetStringValue(SqlDataReader reader, string fieldName, string defaultValue)
        {
            if (reader[fieldName] != DBNull.Value)
            {
                return (string)reader[fieldName];
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets the int value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the integer value</returns>
        public static int GetIntValue(DataRowView row, string fieldName)
        {
            return GetIntValue(row, fieldName, 0);
        }

        /// <summary>
        /// Gets the int value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the integer value</returns>
        public static int GetIntValue(SqlDataReader reader, string fieldName)
        {
            return GetIntValue(reader, fieldName, 0);
        }

        /// <summary>
        /// Gets the int value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns>Returns the integer value</returns>
        public static int GetIntValue(DataRowView row, string fieldName, int defaultValue)
        {
            if (row[fieldName] != DBNull.Value)
            {
                return (int)row[fieldName];
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets the int value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns>Returns the integer value</returns>
        public static int GetIntValue(SqlDataReader reader, string fieldName, int defaultValue)
        {
            if (reader[fieldName] != DBNull.Value)
            {
                return (int)reader[fieldName];
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets the long value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the long value</returns>
        public static long GetLongValue(DataRowView row, string fieldName)
        {
            return GetLongValue(row, fieldName, 0);
        }

        /// <summary>
        /// Gets the long value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the long value</returns>
        public static long GetLongValue(SqlDataReader reader, string fieldName)
        {
            return GetLongValue(reader, fieldName, 0);
        }

        /// <summary>
        /// Gets the long value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns>Returns the long value</returns>
        public static long GetLongValue(DataRowView row, string fieldName, long defaultValue)
        {
            if (row[fieldName] != DBNull.Value)
            {
                return (long)row[fieldName];
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets the long value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns>Returns the long value</returns>
        public static long GetLongValue(SqlDataReader reader, string fieldName, long defaultValue)
        {
            if (reader[fieldName] != DBNull.Value)
            {
                return (long)reader[fieldName];
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets the bool value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the boolean value</returns>
        public static bool GetBoolValue(DataRowView row, string fieldName)
        {
            return GetBoolValue(row, fieldName, false);
        }

        /// <summary>
        /// Gets the bool value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the boolean value</returns>
        public static bool GetBoolValue(SqlDataReader reader, string fieldName)
        {
            return GetBoolValue(reader, fieldName, false);
        }

        /// <summary>
        /// Gets the bool value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">if set to <c>true</c> [default value].</param>
        /// <returns>Returns the boolean value</returns>
        public static bool GetBoolValue(DataRowView row, string fieldName, bool defaultValue)
        {
            if (row[fieldName] != DBNull.Value)
            {
                return (bool)row[fieldName];
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets the bool value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">if set to <c>true</c> [default value].</param>
        /// <returns>Returns the boolean value</returns>
        public static bool GetBoolValue(SqlDataReader reader, string fieldName, bool defaultValue)
        {
            if (reader[fieldName] != DBNull.Value)
            {
                return (bool)reader[fieldName];
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets the date time value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the datetime value</returns>
        public static DateTime GetDateTimeValue(DataRowView row, string fieldName)
        {
            return GetDateTimeValue(row, fieldName, DateTime.MinValue);
        }

        /// <summary>
        /// Gets the date time value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the datetime value</returns>
        public static DateTime GetDateTimeValue(SqlDataReader reader, string fieldName)
        {
            return GetDateTimeValue(reader, fieldName, DateTime.MinValue);
        }

        /// <summary>
        /// Gets the date time value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns>Returns the datetime value</returns>
        public static DateTime GetDateTimeValue(DataRowView row, string fieldName, DateTime defaultValue)
        {
            if (row[fieldName] != DBNull.Value)
            {
                return (DateTime)row[fieldName];
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets the date time value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns>Returns the datetime value</returns>
        public static DateTime GetDateTimeValue(SqlDataReader reader, string fieldName, DateTime defaultValue)
        {
            if (reader[fieldName] != DBNull.Value)
            {
                return (DateTime)reader[fieldName];
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets the decimal value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the decimal value</returns>
        public static decimal GetDecimalValue(DataRowView row, string fieldName)
        {
            return GetDecimalValue(row, fieldName, 0);
        }

        /// <summary>
        /// Gets the decimal value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns>Returns the decimal value</returns>
        public static decimal GetDecimalValue(SqlDataReader reader, string fieldName)
        {
            return GetDecimalValue(reader, fieldName, 0);
        }

        /// <summary>
        /// Gets the decimal value.
        /// </summary>
        /// <param name="row">The data row.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns>Returns the decimal value</returns>
        public static decimal GetDecimalValue(DataRowView row, string fieldName, decimal defaultValue)
        {
            if (row[fieldName] != DBNull.Value)
            {
                return Convert.ToDecimal(row[fieldName]);
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets the decimal value.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns>Returns the decimal value</returns>
        public static decimal GetDecimalValue(SqlDataReader reader, string fieldName, decimal defaultValue)
        {
            if (reader[fieldName] != DBNull.Value)
            {
                return Convert.ToDecimal(reader[fieldName]);
            }

            return defaultValue;
        }

        /// <summary>
        /// Determines whether [is valid var char] [the specified varcharstring].
        /// </summary>
        /// <param name="varcharstring">The varcharstring.</param>
        /// <param name="stringlength">The stringlength.</param>
        /// <returns>Returns true if valid var char</returns>
        public static bool IsValidVarChar(string varcharstring, int stringlength)
        {
            return varcharstring.Length <= stringlength;
        }

        /// <summary>
        /// Determines whether [is valid char] [the specified charstring].
        /// </summary>
        /// <param name="charstring">The charstring.</param>
        /// <param name="stringlength">The stringlength.</param>
        /// <returns>Returns true if valid char</returns>
        public static bool IsValidChar(string charstring, int stringlength)
        {
            return charstring.Length <= stringlength;
        }

        /// <summary>
        /// Method used to finalize a SqlConnection.  This method will Close, Dispose, and destroy
        /// an instance of the SqlConnection passed in.
        /// </summary>
        /// <param name="conn">
        /// The SqlConnection to finalize
        /// </param>
        public void FinalizeConnection(ref SqlConnection conn)
        {
            if (conn != null)
            {
                if (conn.State != ConnectionState.Closed)
                {
                    conn.Close();
                }

                conn.Dispose();
                if (conn == this.connection)
                {
                    this.connection = null;
                }
            }

            conn = null;
        }

        /// <summary>
        /// Method used to Finalize a SqlTransaction.  The method either commits or
        /// rolls a transaction back.
        /// </summary>
        /// <param name="trans">
        /// The SqlTransaction to finalize.
        /// </param>
        /// <param name="commit">
        /// A boolean value indicating whether or not the 
        /// transaction should be committed (true) or rolled back (false).
        /// </param>
        public void FinalizeTransaction(ref SqlTransaction trans, bool commit)
        {
            if (trans != null)
            {
                if (trans.Connection != null)
                {
                    if (commit)
                    {
                        trans.Commit();
                    }
                    else
                    {
                        trans.Rollback();
                    }

                    if (this.transaction == trans)
                    {
                        this.transaction = null;
                    }
                }
            }
        }

        /// <summary>
        /// The purpose of this function is to allow for the execution of
        /// a stored procedure designed to return data.  It should be used for all get and
        /// list stored procedures.
        /// This method supports transactions, if the connection and transaction objects
        /// are passed into the method by reference.  It assumes an Isolation level of ReadUncommitted.
        /// </summary>
        /// <param name="procName">
        /// The name of the stored procedure.
        /// </param>
        /// <param name="parameters">
        /// An array list of parameters to be inserted into the stored procedure.
        /// </param>
        /// <param name="returnTableName">
        /// The name of the table in the dataset created by the stored procedure.
        /// </param>
        /// <param name="conn">
        /// A connection object that has a transaction associated with it.  A null
        /// object may be passed in the first call of a transaction.  Once opened, this connection
        /// will stay opened until the transaction is committed, or an exception occurs.
        /// </param>
        /// <param name="trans">
        /// A transaction object to be used for rollbacks and commits.
        /// </param>
        /// <returns>
        /// An ADO.NET data set containing the data specified in the stored procedure.
        /// </returns>
        public DataSet RunSPReturnData(string procName, ref ArrayList parameters, string returnTableName, ref SqlConnection conn, ref SqlTransaction trans)
        {
            return this.RunSPReturnDataTransaction(procName, ref parameters, returnTableName, ref conn, ref trans, this.isoLevel);
        }

        /// <summary>
        /// Returns a SqlDataReader from a stored procedure call as part of a transaction.  The SqlDataReader's
        /// connection will remain open until the connection is closed and the transaction is either committed or 
        /// rolled back.  Failure to do so will leave the database connection open.  It is the responsibility of the 
        /// consumer of this function to close the database connection.
        /// </summary>
        /// <param name="procName">
        /// The name of the stored procedure.
        /// </param>
        /// <param name="parameters">
        /// An array list of parameters to be inserted into the stored procedure.
        /// </param>
        /// <param name="conn">
        /// A connection object that has a transaction associated with it.  A null
        /// object may be passed in the first call of a transaction.  Once opened, this connection
        /// will stay opened until the transaction is committed, or an exception occurs.
        /// </param>
        /// <param name="trans">
        /// A transaction object to be used for rollbacks and commits.
        /// </param>
        /// <returns>
        /// An ADO.NET SqlDataReader containing the data specified in the stored procedure.
        /// </returns>
        public SqlDataReader RunSPReturnDataReader(string procName, ref ArrayList parameters, ref SqlConnection conn, ref SqlTransaction trans)
        {
            return this.RunSPReturnDataReader(procName, ref parameters, ref conn, ref trans, this.isoLevel);
        }

        /// <summary>
        /// Returns a SqlDataReader from a stored procedure call as part of a transaction.  The SqlDataReader's
        /// connection will remain open until the connection is closed and the transaction is either committed or 
        /// rolled back.  Failure to do so will leave the database connection open.  It is the responsibility of the 
        /// consumer of this function to close the database connection.
        /// </summary>
        /// <param name="procName">
        /// The name of the stored procedure.
        /// </param>
        /// <param name="parameters">
        /// An array list of parameters to be inserted into the stored procedure.
        /// </param>
        /// <param name="conn">
        /// A connection object that has a transaction associated with it.  A null
        /// object may be passed in the first call of a transaction.  Once opened, this connection
        /// will stay opened until the transaction is committed, or an exception occurs.
        /// </param>
        /// <param name="trans">
        /// A transaction object to be used for rollbacks and commits.
        /// </param>
        /// <param name="transIsoLevel">
        /// The transaction isolation level. Ex: ReadCommitted
        /// </param>
        /// <returns>
        /// An ADO.NET SqlDataReader containing the data specified in the stored procedure.
        /// </returns>
        public SqlDataReader RunSPReturnDataReader(string procName, ref ArrayList parameters, ref SqlConnection conn, ref SqlTransaction trans, IsolationLevel transIsoLevel)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command
            bool closeConnection = false;

            try
            {
                this.InitializeConnection(ref conn, ref trans, transIsoLevel);

                // Set command information
                command = new SqlCommand
                              {
                                  CommandType = CommandType.StoredProcedure,
                                  CommandText = procName,
                                  Connection = this.connection,
                                  CommandTimeout = this.CommandTiemout,
                                  Transaction = this.transaction
                              };
                AddParametersToCommand(ref command, parameters);

                // Go get the data, and return the datareader
                var reader = command.ExecuteReader();
                parameters = FillOutputParameters(command, parameters);
                return reader;
            }
            catch (Exception)
            {
                closeConnection = true; // Always close if there is an error
                throw;
            }
            finally
            {
                // Destroy objects
                if (command != null)
                {
                    if (closeConnection)
                    {
                        command.Transaction.Rollback();
                    }

                    command.Dispose();
                    command = null;
                }

                if (this.connection != null && closeConnection)
                {
                    this.connection.Close();
                    this.connection.Dispose();
                    this.connection = null;
                    this.transaction = null;
                    conn = null;
                    trans = null;
                }
            }
        }

        /// <summary>
        /// To provide a way to execute a stored procedure, and receive a return value
        /// from the stored procedure.  This stored procedure will primarily be used for insert
        /// stored procedures that would create an identity value for a primary key, and then
        /// return that value through the stored procedure.
        /// The method supports transactions, and the isolation level for this method is ReadUncommitted.
        /// </summary>
        /// <param name="procName">Name of the proc.</param>
        /// <param name="retParamName">Name of the ret param.</param>
        /// <param name="parameters">The parameters.</param>
        /// <param name="conn">The connection.</param>
        /// <param name="trans">The transaction.</param>
        /// <returns>The run sp return number.</returns>
        public int RunSPReturnNumber(string procName, string retParamName, ref ArrayList parameters, ref SqlConnection conn, ref SqlTransaction trans)
        {
            return this.RunSPReturnNumberTransaction(procName, retParamName, ref parameters, ref conn, ref trans, this.isoLevel);
        }

        /// <summary>
        /// Provides a mean to run a stored procedure used to perform execute 
        /// commands.  This stored procedure should be used for update, delete, and insert commands
        /// that do not require a return value such as an identity.
        /// At the moment, this method does NOT support transactions.
        /// </summary>
        /// <param name="procName">
        /// The name of the stored procedure.
        /// </param>
        /// <param name="parameters">
        /// A list of parameters to be inserted into the stored procedure.
        /// </param>
        /// <returns>
        /// The number of records affected by the stored procedure.
        /// </returns>
        public int RunSP(string procName, ref ArrayList parameters)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command
            bool useGlobal = false;

            try
            {
                useGlobal = this.InitializeConnection();

                // Set command information
                command = new SqlCommand
                              {
                                  CommandType = CommandType.StoredProcedure,
                                  CommandText = procName,
                                  Connection = this.connection,
                                  CommandTimeout = this.CommandTiemout
                              };
                command.CommandTimeout = this.CommandTiemout;
                if (this.transaction != null)
                {
                    command.Transaction = this.transaction;
                }

                AddParametersToCommand(ref command, parameters);

                // Go get the data, and return the number of records affected
                int recsAffected = command.ExecuteNonQuery();
                parameters = FillOutputParameters(command, parameters);
                return recsAffected;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                return 0;
            }
            finally
            {
                // Destroy objects
                if (!useGlobal)
                {
                    this.FinalizeConnection(ref this.connection);
                }

                if (command != null)
                {
                    command.Dispose();
                }
            }
        }

        /// <summary>
        /// Provides a mean to run a stored procedure used to perform execute 
        /// commands.  This stored procedure should be used for update, delete, and insert commands
        /// that do not require a return value such as an identity.  The isolation level for this method
        /// is ReadUncommitted
        /// </summary>
        /// <param name="procName">
        /// The name of the stored procedure.
        /// </param>
        /// <param name="parameters">
        /// A list of parameters to be inserted into the stored procedure.
        /// </param>
        /// <param name="conn">
        /// A connection object that has a transaction associated with it.  A null
        /// object may be passed in the first call of a transaction.  Once opened, this connection
        /// will stay opened until the transaction is committed, or an exception occurs.
        /// </param>
        /// <param name="trans">
        /// A transaction object to be used for rollbacks and commits.
        /// </param>
        /// <returns>
        /// The run sp.
        /// </returns>
        public int RunSP(string procName, ref ArrayList parameters, ref SqlConnection conn, ref SqlTransaction trans)
        {
            return this.RunSPTransaction(procName, ref parameters, ref conn, ref trans, this.isoLevel);
        }

        /// <summary>
        /// The purpose of this function is to allow for the execution of
        /// a stored procedure designed to return data.  It should be used for all get and
        /// list stored procedures.
        /// At the moment, this method does NOT support transactions.
        /// </summary>
        /// <param name="procName">
        /// The name of the stored procedure.
        /// </param>
        /// <param name="parameters">
        /// An array list of parameters to be inserted into the stored procedure.
        /// </param>
        /// <param name="returnTableName">
        /// The name of the table in the dataset created by the stored procedure
        /// </param>
        /// <returns>
        /// An ADO.NET data set containing the data specified in the stored procedure.
        /// </returns>
        public DataSet RunSPReturnData(string procName, ref ArrayList parameters, string returnTableName)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command
            SqlDataAdapter adapter = null;
            DataSet dataSet = null;
            bool useGlobal = false;

            try
            {
                // First, create instances of objects
                useGlobal = this.InitializeConnection();
                adapter = new SqlDataAdapter();
                dataSet = new DataSet();

                // Set command information
                command = new SqlCommand { Connection = this.connection };
                if (this.transaction != null)
                {
                    command.Transaction = this.transaction;
                }

                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = procName;
                command.CommandTimeout = this.CommandTiemout;
                AddParametersToCommand(ref command, parameters);

                // Set adapter information
                adapter.SelectCommand = command;

                // Go get the data, and return the dataset
                adapter.Fill(dataSet, returnTableName);
                parameters = FillOutputParameters(adapter.SelectCommand, parameters);
                return dataSet;
            }
            finally
            {
                // Destroy objects
                if (!useGlobal)
                {
                    this.FinalizeConnection(ref this.connection);
                }

                if (command != null)
                {
                    command.Dispose();
                    command = null;
                }

                if (adapter != null)
                {
                    adapter.Dispose();
                    adapter = null;
                }

                if (dataSet != null)
                {
                    dataSet.Dispose();
                    dataSet = null;
                }
            }
        }

        /// <summary>
        /// Method used to obtain a SqlDataReader object by calling a stored procedure from a 
        /// given set of parameters.  This method takes care of obtaining the connection, and calling
        /// the stored procedure specified.  IMPORTANT: Due to the nature of the SqlDataReader object,
        /// this method does NOT close the connection.  The connection will close when the developer using
        /// this object calls the SqlDataReader.Close() method within the finally block of the calling function.
        /// If you wish to take advantage of connection pooling, you must call the Close method on the SqlDataReader
        /// object that is returned.
        /// </summary>
        /// <param name="procName">
        /// The name of the stored procedure.
        /// </param>
        /// <param name="parameters">
        /// An array list of parameter structs to be
        /// added to the stored procedure.
        /// </param>
        /// <returns>
        /// An SqlDataReader object populated with specified data.
        /// </returns>
        public SqlDataReader RunSPReturnDataReader(string procName, ref ArrayList parameters)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command

            try
            {
                bool useGlobal = this.InitializeConnection();

                command = new SqlCommand { Connection = this.connection };

                // Set command information
                if (this.transaction != null)
                {
                    command.Transaction = this.transaction;
                }

                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = procName;
                command.CommandTimeout = this.CommandTiemout;
                AddParametersToCommand(ref command, parameters);

                SqlDataReader reader = useGlobal ? command.ExecuteReader() : command.ExecuteReader(CommandBehavior.CloseConnection);

                parameters = FillOutputParameters(command, parameters);
                return reader;
            }
            finally
            {
                // Destroy objects
                if (command != null)
                {
                    command.Dispose();
                    command = null;
                }
            }
        }
 
        /// <summary>
        /// To provide a way to execute a stored procedure, and receive a return value
        /// from the stored procedure.  This stored procedure will primarily be used for insert
        /// stored procedures that would create an identity value for a primary key, and then
        /// return that value through the stored procedure.
        /// </summary>
        /// <param name="procName">
        /// The name of the stored procedure to execute.
        /// </param>
        /// <param name="retParamName">
        /// The name of the stored procedure return value.
        /// </param>
        /// <param name="parameters">
        /// A list of parameters to be inserted into the stored procedure.
        /// </param>
        /// <returns>
        /// The value returned by the stored procedure, specified by the value passed into the retParamName parameter.
        /// </returns>
        public int RunSPReturnNumber(string procName, string retParamName, ref ArrayList parameters)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command
            bool useGlobal = false;

            try
            {
                useGlobal = this.InitializeConnection();

                // Set command information
                command = new SqlCommand
                              {
                                  CommandType = CommandType.StoredProcedure,
                                  CommandText = procName,
                                  Connection = this.connection,
                                  CommandTimeout = this.CommandTiemout
                              };
                if (this.transaction != null)
                {
                    command.Transaction = this.transaction;
                }

                // Add the parameter to the list
                parameters.Add(new Parameter(retParamName, DataTypes.IntDT, ParamDirections.OutputPD, null, 4));

                AddParametersToCommand(ref command, parameters);

                // Execute the stored procedure
                command.ExecuteNonQuery();

                // Now return the return value
                parameters = FillOutputParameters(command, parameters);

                if (command.Parameters[retParamName].Value != DBNull.Value)
                {
                    return (int) command.Parameters[retParamName].Value;
                }
                else
                {
                    throw new Exception("A null value cannot be returned as an output parameter on a stored procedure (" + procName + ")");
                }
            }
            finally
            {
                // Destroy objects
                if (!useGlobal)
                {
                    this.FinalizeConnection(ref this.connection);
                }

                if (command != null)
                {
                    command.Dispose();
                    command = null;
                }
            }
        }
 
        /// <summary>
        /// Function used to run a given sql string, execute the sql command, and return a single value
        /// It should be used for commands that bring back a singel value. 
        /// Since this function returns a System.Object, you will need to cast the
        /// return value as whatever you were expecting.
        /// At the moment, this method does  NOT support transactions.
        /// </summary>
        /// <param name="sql">
        /// The sql string that is to be run against the database
        /// </param>
        /// <returns>
        /// The number of records affected by the operation
        /// </returns>
        public object RunSQLExecuteScalar(string sql)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command
            bool useGlobal = false;

            try
            {
                useGlobal = this.InitializeConnection();

                // Set command information
                command = new SqlCommand
                              {
                                  CommandType = CommandType.Text,
                                  CommandText = sql,
                                  Connection = this.connection,
                                  CommandTimeout = this.CommandTiemout
                              };
                if (this.transaction != null)
                {
                    command.Transaction = this.transaction;
                }

                // Go get the data, and return the number of records affected
                return command.ExecuteScalar();
            }
            finally
            {
                // Destroy objects
                if (!useGlobal)
                {
                    this.FinalizeConnection(ref this.connection);
                }

                if (command != null)
                {
                    command.Dispose();
                    command = null;
                }
            }
        }
 
        /// <summary>
        /// Function used to run a given sql string, and execute the sql command.
        /// It should be used for Insert, Update, and Delete commands.  At the moment, this method does
        /// NOT support transactions.
        /// </summary>
        /// <param name="sql">
        /// The sql string that is to be run against the database
        /// </param>
        /// <returns>
        /// The number of records affected by the operation
        /// </returns>
        public int RunSQLExecute(string sql)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command
            bool useGlobal = false;

            try
            {
                useGlobal = this.InitializeConnection();

                // Set command information
                command = new SqlCommand
                              {
                                  CommandType = CommandType.Text,
                                  CommandText = sql,
                                  Connection = this.connection,
                                  CommandTimeout = this.CommandTiemout
                              };
                if (this.transaction != null)
                {
                    command.Transaction = this.transaction;
                }

                return command.ExecuteNonQuery();
            }
            finally
            {
                // Destroy objects
                if (!useGlobal)
                {
                    this.FinalizeConnection(ref this.connection);
                }

                if (command != null)
                {
                    command.Dispose();
                    command = null;
                }
            }
        }
 
        /// <summary>
        /// Function used to run a given sql string, and return an ADO.NET dataset.
        /// At the moment, this method does NOT support transactions.
        /// </summary>
        /// <param name="sql">
        /// The sql string that is to be run against the database
        /// </param>
        /// <param name="returnTableName">
        /// The name of the table in the dataset created by the sql operation
        /// </param>
        /// <returns>
        /// ADO.NET dataset containing requested data
        /// </returns>
        public DataSet RunSQLReturnData(string sql, string returnTableName)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command
            SqlDataAdapter adapter = null;
            DataSet dataSet = null;
            bool useGlobal = false;

            try
            {
                // First, create instances of objects
                useGlobal = this.InitializeConnection();
                adapter = new SqlDataAdapter();
                dataSet = new DataSet();

                // Set command information
                command = new SqlCommand
                              {
                                  CommandType = CommandType.Text,
                                  CommandText = sql,
                                  Connection = this.connection,
                                  CommandTimeout = this.CommandTiemout
                              };
                if (this.transaction != null)
                {
                    command.Transaction = this.transaction;
                }

                // Set adapter information
                adapter.SelectCommand = command;

                // Go get the data, and return the dataset
                adapter.Fill(dataSet, returnTableName);
                return dataSet;
            }
            finally
            {
                // Destroy objects
                if (!useGlobal)
                {
                    this.FinalizeConnection(ref this.connection);
                }

                if (command != null)
                {
                    command.Dispose();
                    command = null;
                }

                if (adapter != null)
                {
                    adapter.Dispose();
                    adapter = null;
                }

                if (dataSet != null)
                {
                    dataSet.Dispose();
                    dataSet = null;
                }
            }
        }
 
        /// <summary>
        /// Method used to obtain a SqlDataReader object by running the provided sql.
        /// This method takes care of obtaining the connection, and running
        /// the sql specified.  IMPORTANT: Due to the nature of the SqlDataReader object,
        /// this method does NOT close the connection.  The connection will close when the developer using
        /// this object calls the SqlDataReader.Close() method within the finally block of the calling function.
        /// If you wish to take advantage of connection pooling, you must call the Close method on the SqlDataReader
        /// object that is returned. 
        /// </summary>
        /// <param name="sql">
        /// The sql string to run.
        /// </param>
        /// <returns>
        /// An SqlDataReader object populated with specified data.
        /// </returns>
        public SqlDataReader RunSQLReturnDataReader(string sql)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command

            try
            {
                // First, create instances of objects
                bool useGlobal = this.InitializeConnection();

                // Set command information
                command = new SqlCommand
                              {
                                  CommandType = CommandType.Text,
                                  CommandText = sql,
                                  Connection = this.connection,
                                  CommandTimeout = this.CommandTiemout
                              };
                if (this.transaction != null)
                {
                    command.Transaction = this.transaction;
                }

                // Go get the data, and return the reader
                // This will not automatically close the connection.  The developer using
                // this method must call SqlDataReader.Close() in order for the connection
                // to actually be closed.
                return useGlobal ? command.ExecuteReader() : command.ExecuteReader(CommandBehavior.CloseConnection);
            }
            finally
            {
                if (command != null)
                {
                    command.Dispose();
                    command = null;
                }
            }
        }

        /// <summary>
        /// Purpose:  This function provides a method of adding the parameters from a 
        /// parameter collection to a command object, which is passed in by reference.<newpara></newpara>
        /// </summary>
        /// <param name="command">
        /// The ADO.NET SqlCommand you would like to insert the parameters into
        /// </param>
        /// <param name="parameters">
        /// The parameters to add to the SqlCommand.
        /// </param>
        private static void AddParametersToCommand(ref SqlCommand command, ArrayList parameters)
        {
            SqlParameter sqlParam = null;

            if (parameters.Count > 0)
            {
                foreach (Parameter param in parameters)
                {
                    // Place each parameter into the command object
                    sqlParam = new SqlParameter { ParameterName = param.ParamName, Value = param.Value };

                    switch (param.DataType)
                    {
                        case DataTypes.CharDT:
                            if (param.Value == null)
                            {
                                sqlParam.Value = DBNull.Value;
                            }

                            break;

                        case DataTypes.VarCharDT:
                            if (param.Value == null)
                            {
                                sqlParam.Value = DBNull.Value;
                            }

                            break;

                        case DataTypes.TextDT:
                            if (param.Value == null)
                            {
                                sqlParam.Value = DBNull.Value;
                            }

                            break;

                        default:

                            sqlParam.Value = param.Value;
                            break;
                    }

                    // Set necessary sql parameter properties
                    sqlParam.SqlDbType = (SqlDbType)param.DataType;
                    sqlParam.Size = param.Size;
                    sqlParam.Direction = (ParameterDirection)param.ParamDirection;

                    // Add to sql command
                    command.Parameters.Add(sqlParam);
                }
            }
        }

        /// <summary>
        /// Method used to fill a list of OUTPUT parameters to send back to the caller.  It is useful
        /// if the stored procedure called has multiple output parameters that need to be returned 
        /// with a dataset, datareader, or some other value.
        /// </summary>
        /// <param name="command">
        /// The SqlCommand object that was used to call a stored procedure.
        /// </param>
        /// <param name="parameters">
        /// The original list of parameters, which is needed by the method to 
        /// know which output parameters in the command object to look for.
        /// </param>
        /// <returns>
        /// A new list of output parameters, with the values created by the stored procedure
        /// inside.
        /// </returns>
        private static ArrayList FillOutputParameters(SqlCommand command, ArrayList parameters)
        {
            ArrayList tempParameters = null;
            try
            {
                if (parameters.Count > 0)
                {
                    tempParameters = new ArrayList();
                    tempParameters = (ArrayList)parameters.Clone(); // Set the temp list of parameters = the original
                    parameters.Clear(); // Clear, so that we can refresh with output parameters.

                    foreach (Parameter param in tempParameters)
                    {
                        Parameter newParam;
                        switch (param.ParamDirection)
                        {
                            case ParamDirections.OutputPD:
                                newParam = new Parameter(param.ParamName, param.DataType, param.ParamDirection, command.Parameters[param.ParamName].Value, param.Size);
                                parameters.Add(newParam);
                                break;
                            case ParamDirections.InputOutputPD:
                                newParam = new Parameter(param.ParamName, param.DataType, param.ParamDirection, command.Parameters[param.ParamName].Value, param.Size);
                                parameters.Add(newParam);
                                break;
                            case ParamDirections.ReturnValuePD:
                                newParam = new Parameter(param.ParamName, param.DataType, param.ParamDirection, command.Parameters[param.ParamName].Value, param.Size);
                                parameters.Add(newParam);
                                break;
                            default:
                                break;
                        }
                    }
                }

                return parameters;
            }
            finally
            {
                tempParameters = null;
            }
        }

        /// <summary>
        /// Retrieve an Key Value pair from the appSettings section of the web.config file.  If the
        /// file parameter of the appSettings tag is set to a valid file reference, this
        /// method will retrieve values from that file.
        /// </summary>
        /// <param name="key">
        /// The key whose value you want to retrieve
        /// </param>
        /// <returns>
        /// The value corresponding to Key
        /// </returns>
        private static string GetConfigElement(string key)
        {
            string val = ConfigurationManager.AppSettings[key];
            if (val == null)
            {
                throw new Exception("Could not retrieve key " + key + " from web.config");
            }

            if (!val.Trim().Equals(string.Empty))
            {
                return val;
            }

            throw new Exception("The Key " + key + " was empty in the configuration file");
        }

        /// <summary>
        /// Method used to set the connection string property of the DataAccess class. If the
        /// CryptoProvider property has been set, this method will use it to decrypt the connection
        /// string.
        /// </summary>
        private void SetConnectionString()
        {
            this.connectionString = GetConfigElement(this.ConnectionStringKey);
        }

        /// <summary>
        /// Configures the 'connection' field and opens it for use.
        /// </summary>
        /// <returns>
        /// True if this is the reused 'connection' field, false if newly created.
        /// </returns>
        private bool InitializeConnection()
        {
            // First, create instances of objects
            if (this.connection != null && this.transaction != null
                && this.connection.State == ConnectionState.Open)
            {
                return true;
            }

            this.connection = new SqlConnection();
            this.transaction = null;

            // Set connection string value
            this.SetConnectionString();
            this.connection.ConnectionString = this.connectionString;
            this.connection.Open();
            return false;
        }

        /// <summary>
        /// Configures the 'connection' field and opens it for use.
        /// </summary>
        /// <param name="conn">A connection object to return.</param>
        /// <param name="trans">A transaction object to return.</param>
        /// <param name="transIsoLevel">The transaction isolation level. Ex: ReadCommitted</param>
        private void InitializeConnection(ref SqlConnection conn, ref SqlTransaction trans, IsolationLevel transIsoLevel)
        {
            // First, create instances of objects
            if (conn != null && conn != this.connection)
            {
                if (this.transaction != null)
                {
                    this.FinalizeTransaction(ref this.transaction, false);
                }

                if (this.connection != null)
                {
                    this.FinalizeConnection(ref this.connection);
                }

                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }

                if (trans == null)
                {
                    trans = conn.BeginTransaction(transIsoLevel);
                }

                this.connection = conn;
                this.transaction = trans;
                return;
            }

            if (this.connection != null && this.transaction != null
                && this.connection.State == ConnectionState.Open)
            {
                conn = this.connection;
                trans = this.transaction;
                return;
            }

            this.connection = new SqlConnection();

            // Set connection string value
            this.SetConnectionString();
            this.connection.ConnectionString = this.connectionString;
            this.connection.Open();
            this.transaction = this.connection.BeginTransaction(transIsoLevel);
            conn = this.connection;
            trans = this.transaction;
            return;
        }

        /// <summary>
        /// Runs a stored procedure within the context of a transaction.
        /// </summary>
        /// <param name="procName">
        /// The stored procedure to be executed.
        /// </param>
        /// <param name="parameters">
        /// The list of paramters to be added to the stored procedure.
        /// </param>
        /// <param name="conn">
        /// The database connection object to be used.
        /// </param>
        /// <param name="trans">
        /// The transaction object to contain the stored procedure call.
        /// </param>
        /// <param name="transIsoLevel">
        /// The transaction isolation level. Ex: ReadCommitted
        /// </param>
        /// <returns>
        /// The number of records affected by the operation.
        /// </returns>
        private int RunSPTransaction(string procName, ref ArrayList parameters, ref SqlConnection conn, ref SqlTransaction trans, IsolationLevel transIsoLevel)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command
            bool closeConnection = false;

            try
            {
                this.InitializeConnection(ref conn, ref trans, transIsoLevel);

                // Set command information
                command = new SqlCommand
                              {
                                  CommandType = CommandType.StoredProcedure,
                                  CommandText = procName,
                                  Connection = this.connection,
                                  CommandTimeout = this.CommandTiemout,
                                  Transaction = this.transaction
                              };
                AddParametersToCommand(ref command, parameters);

                // Go get the data, and return the number of records affected
                int recsAffected = command.ExecuteNonQuery();

                parameters = FillOutputParameters(command, parameters);
                return recsAffected;
            }
            catch (Exception)
            {
                closeConnection = true; // Always close if there is an error
                throw;
            }
            finally
            {
                // Destroy objects
                if (command != null)
                {
                    if (closeConnection)
                    {
                        command.Transaction.Rollback();
                    }

                    command.Dispose();
                    command = null;
                }

                if (this.connection != null && closeConnection)
                {
                    this.connection.Close();
                    this.connection.Dispose();
                    this.connection = null;
                    this.transaction = null;
                    conn = null;
                    trans = null;
                }
            }
        }

        /// <summary>
        /// Runs a stored procedure within the context of a transaction, and also returns the value of the specified retParamName.
        /// </summary>
        /// <param name="procName">The stored procedure to be executed.</param>
        /// <param name="retParamName">The name of the output parameter in the stored procedure.</param>
        /// <param name="parameters">The list of paramters to be added to the stored procedure.</param>
        /// <param name="conn">The database connection object to be used.</param>
        /// <param name="trans">The transaction object to contain the stored procedure call.</param>
        /// <param name="transIsoLevel">The transaction isolation level. Ex: ReadCommitted</param>
        /// <returns>
        /// The value of the parameter specified by retParamName
        /// </returns>
        private int RunSPReturnNumberTransaction(string procName, string retParamName, ref ArrayList parameters, ref SqlConnection conn, ref SqlTransaction trans, IsolationLevel transIsoLevel)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command
            bool closeConnection = false;
            try
            {
                this.InitializeConnection(ref conn, ref trans, transIsoLevel);

                // Set command information
                command = new SqlCommand
                              {
                                  CommandType = CommandType.StoredProcedure,
                                  CommandText = procName,
                                  Connection = this.connection,
                                  CommandTimeout = this.CommandTiemout,
                                  Transaction = this.transaction
                              };

                // Create the output return parameter
                parameters.Add(new Parameter(retParamName, DataTypes.IntDT, ParamDirections.OutputPD, null, 4));
                AddParametersToCommand(ref command, parameters);

                // Go get the data, and return the number of records affected
                command.ExecuteNonQuery();

                // Now return the return value
                parameters = FillOutputParameters(command, parameters);

                if (command.Parameters[retParamName].Value != DBNull.Value)
                {
                    return (int)command.Parameters[retParamName].Value;
                }
                else
                {
                    throw new Exception("A null value cannot be returned as an output parameter on a stored procedure (" + procName + ")");
                }
            }
            catch (Exception)
            {
                closeConnection = true; // Always close if there is an error
                throw;
            }
            finally
            {
                // Destroy objects
                if (command != null)
                {
                    if (closeConnection)
                    {
                        command.Transaction.Rollback();
                    }

                    command.Dispose();
                    command = null;
                }

                if (this.connection != null && closeConnection)
                {
                    this.connection.Close();
                    this.connection.Dispose();
                    this.connection = null;
                    this.transaction = null;
                    conn = null;
                    trans = null;
                }
            }
        }

        /// <summary>
        /// Runs a stored procedure that returns a dataset within the context of a transaction.
        /// </summary>
        /// <param name="procName">
        /// The stored procedure to be executed.
        /// </param>
        /// <param name="parameters">
        /// The list of paramters to be added to the stored procedure.
        /// </param>
        /// <param name="returnTableName">
        /// The desired name of the Dataset table.
        /// </param>
        /// <param name="conn">
        /// The database connection object to be used.
        /// </param>
        /// <param name="trans">
        /// The transaction object to contain the stored procedure call.
        /// </param>
        /// <param name="transIsoLevel">
        /// The transaction isolation level. Ex: ReadCommitted
        /// </param>
        /// <returns>
        /// A disconnected ADO.NET dataset, which is a result of the stored procedure.
        /// </returns>
        private DataSet RunSPReturnDataTransaction(string procName, ref ArrayList parameters, string returnTableName, ref SqlConnection conn, ref SqlTransaction trans, IsolationLevel transIsoLevel)
        {
            // Declaring variables
            SqlCommand command = null; // Object representing the sql command
            SqlDataAdapter adapter = null;
            DataSet dataSet = null;
            bool closeConnection = false;

            try
            {
                this.InitializeConnection(ref conn, ref trans, transIsoLevel);

                // First, create instances of objects
                command = new SqlCommand();
                adapter = new SqlDataAdapter();
                dataSet = new DataSet();

                // Set command information
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = procName;
                command.Connection = this.connection;
                command.CommandTimeout = this.CommandTiemout;
                command.Transaction = this.transaction;
                AddParametersToCommand(ref command, parameters);

                // Set adapter information
                adapter.SelectCommand = command;

                // Go get the data
                adapter.Fill(dataSet, returnTableName);
                parameters = FillOutputParameters(adapter.SelectCommand, parameters);
                return dataSet; // Return the dataset
            }
            catch (Exception)
            {
                closeConnection = true; // Always close if there is an error
                throw;
            }
            finally
            {
                // Destroy objects
                if (command != null)
                {
                    if (closeConnection)
                    {
                        command.Transaction.Rollback();
                    }

                    command.Dispose();
                    command = null;
                }

                if (this.connection != null && closeConnection)
                {
                    this.connection.Close();
                    this.connection.Dispose();
                    this.connection = null;
                    this.transaction = null;
                    conn = null;
                    trans = null;
                }

                if (adapter != null)
                {
                    adapter.Dispose();
                    adapter = null;
                }

                if (dataSet != null)
                {
                    dataSet.Dispose();
                    dataSet = null;
                }
            }
        }
        #endregion Methods

        #region Structs

        /// <summary>
        /// The parameter.
        /// </summary>
        public struct Parameter
        {
            /// <summary>
            /// The data type of the parameter ex: VarCharDT, IntDT
            /// </summary>
            public DataTypes DataType;

            /// <summary>
            /// The direction of the parameter. ex:  Input, Output
            /// </summary>
            public ParamDirections ParamDirection;

            /// <summary>
            /// The name of the parameter to match the parameter name in the stored procedure.
            /// </summary>
            public string ParamName;

            /// <summary>
            /// The size of the datatype.  Ex:  A VARCHAR parameter that has a length of 255 should
            /// have a size of 255.
            /// </summary>
            public int Size;

            /// <summary>
            /// The actual data value of the parameter.
            /// </summary>
            public object Value;

            /// <summary>
            /// Initializes a new instance of the <see cref="Parameter"/> struct. 
            /// Constructor for the Parameter struct
            /// </summary>
            /// <param name="paramName">
            /// The name of the parameter to match the parameter name in the stored procedure.
            /// </param>
            /// <param name="dataType">
            /// The data type of the parameter ex: VarCharDT, IntDT
            /// </param>
            /// <param name="paramDirection">
            /// The direction of the parameter. ex:  Input, Output
            /// </param>
            /// <param name="paramValue">
            /// The actual data value of the parameter.
            /// </param>
            /// <param name="size">
            /// The size of the datatype.  Ex:  A VARCHAR parameter that has a length of 255 should
            /// have a size of 255.
            /// </param>
            public Parameter(string paramName, DataTypes dataType, ParamDirections paramDirection, object paramValue, int size)
            {
                this.ParamName = paramName;
                this.DataType = dataType;
                this.ParamDirection = paramDirection;
                this.Value = paramValue;
                this.Size = size;
            }
        }

        #endregion //Structs
    } 
} 