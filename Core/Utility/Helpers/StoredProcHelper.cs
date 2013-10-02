// --------------------------------------------------------------------------------------------------------------------
// <copyright file="StoredProcHelper.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the StoredProcHelper type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.Utility.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Data.SqlClient;

    /// <summary>
    /// Helps to create a string to call stored procedures in a SAP execute statement
    /// </summary>
    public class StoredProcHelper
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StoredProcHelper"/> class.
        /// </summary>
        /// <param name="procedureName">Name of the stored procedure to call.</param>
        public StoredProcHelper(string procedureName)
        {
            this.ProcedureName = procedureName;
            this.Parameters = new List<SqlParameter>();
        }

        /// <summary>
        /// Gets or sets a value indicating whether to replace the single quotes with double single quotes.
        /// </summary>
        /// <value>
        ///     <c>true</c> if [ignore single quote replace]; otherwise, <c>false</c>.
        /// </value>
        public bool IgnoreSingleQuoteReplace { get; set; }

        /// <summary>
        /// Gets or sets the name of the procedure.
        /// </summary>
        /// <value>The name of the procedure.</value>
        public string ProcedureName { get; set; }

        /// <summary>
        /// Gets or sets the parameters.
        /// </summary>
        /// <value>The parameters.</value>
        public List<SqlParameter> Parameters { get; set; }

        /// <summary>
        /// Adds the parameter.
        /// </summary>
        /// <param name="parameterName">Name of the parameter.</param>
        /// <param name="parameterType">Type of the parameter.</param>
        /// <param name="parameterValue">The parameter value.</param>
        public void AddParameter(string parameterName, System.Data.DbType parameterType, object parameterValue)
        {
            if (!parameterName.StartsWith("@"))
            {
                throw new Exception("A parameter name must start with '@'");
            }

            this.Parameters.Add(new SqlParameter(parameterName, parameterValue));
            this.Parameters[this.Parameters.Count - 1].DbType = parameterType;
        }

        /// <summary>
        /// Returns a <see cref="System.String"/> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String"/> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            if (this.ProcedureName == null)
            {
                throw new NullReferenceException("Procedure Name is missing");
            }

            string queryString = "exec {0} {1}";

            queryString = string.Format(queryString, this.ProcedureName, this.GetParameterString()).Trim(' ');

            return queryString;
        }

        /// <summary>
        /// Gets the parameter string.
        /// </summary>
        /// <returns>Returns a string containing all the parameters formatted for an exec statement</returns>
        private string GetParameterString()
        {
            const string Single_param_alpha = "{0}='{1}', ";
            const string Single_param = "{0}={1}, ";
            string parameterString = string.Empty;
            foreach (SqlParameter param in this.Parameters)
            {
                if (param.Value == null)
                {
                    parameterString += string.Format(Single_param, param.ParameterName, "NULL");
                }
                else
                {
                    string parValue = this.IgnoreSingleQuoteReplace ? param.Value.ToString() : param.Value.ToString().Replace("'", "''");

                    switch (param.DbType)
                    {
                        case System.Data.DbType.String:
                        case System.Data.DbType.StringFixedLength:
                            parameterString += string.Format(Single_param_alpha, param.ParameterName, parValue);
                            break;
                        case System.Data.DbType.Time:
                        case System.Data.DbType.Date:
                        case System.Data.DbType.DateTime:
                        case System.Data.DbType.DateTime2:
                            parameterString += string.Format(Single_param_alpha, param.ParameterName, parValue);
                            break;
                        default:
                            parameterString += string.Format(Single_param, param.ParameterName, parValue);
                            break;
                    }
                }
            }

            if (parameterString.EndsWith(", "))
            {
                parameterString = parameterString.Substring(0, (parameterString.Length - 2));
            }

            return parameterString;
        }
    }
}
