// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SqlHelper.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   This class builds parametarized Sql statements
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.Utility.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Data.SqlClient;
    using System.Text;
    
    /// <summary>
    /// This class builds parametarized Sql statements
    /// </summary>
    public class SqlHelper
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SqlHelper"/> class.
        /// </summary>
        public SqlHelper()
        {
            this.Parameters = new List<SqlParameter>();
            this.Builder = new StringBuilder();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlHelper"/> class.
        /// </summary>
        /// <param name="sqlStatement">The SQL statement.</param>
        public SqlHelper(string sqlStatement) : this()
        {
            this.SqlStatement = sqlStatement;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlHelper"/> class.
        /// </summary>
        /// <param name="sqlStatementBuilder">The SQL statement builder.</param>
        public SqlHelper(StringBuilder sqlStatementBuilder) : this(sqlStatementBuilder.ToString())           
        {
            this.Builder = sqlStatementBuilder;
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
        public string SqlStatement { get; set; }

        /// <summary>
        /// Gets or sets the builder.
        /// </summary>
        /// <value>The builder.</value>
        public StringBuilder Builder { get; set; }

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

            if (string.IsNullOrEmpty(this.SqlStatement))
            {
                this.SqlStatement = this.Builder.ToString();
            }

            if (!this.SqlStatement.Contains(parameterName))
            {
                throw new Exception("The parameter is not contained in the Sql Statement.");
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
            if (!string.IsNullOrEmpty(this.Builder.ToString()))
            {
                this.SqlStatement = this.Builder.ToString();
            }

            if (string.IsNullOrEmpty(this.SqlStatement))
            {
                throw new NullReferenceException("Sql Statement is missing");
            }

            string queryString = "exec sp_executesql{0}N'{1}',{0} N'{2}',{0}{3}";

            // {0} = Environment New Line
            // {1} = SQL Statement
            // {2} = Parameter declaration
            // {3} = Parameters list
            queryString = string.Format(
                queryString, 
                Environment.NewLine, 
                this.SqlStatement.Replace("'", "''"), 
                this.GetParameterDeclarationString(),
                this.GetParameterString()).Trim(' ');

            if (queryString.EndsWith("," + Environment.NewLine))
            {
                queryString = queryString.Substring(0, (queryString.Length - 3));
            }

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

                            // The Recorset object passes a value of '31/12/1899' when the DateTime is null.
                            // We need to convert the value back to null in this case.
                            DateTime dt;
                            DateTime.TryParse(parValue, out dt);
                            parameterString += dt.Year < 1900
                                                   ? string.Format(Single_param, param.ParameterName, "NULL")
                                                   : string.Format(Single_param_alpha, param.ParameterName, parValue);

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

            parameterString.Replace(", ", "," + Environment.NewLine);

            return parameterString;
        }

        /// <summary>
        /// Gets the parameter declaration string.
        /// </summary>
        /// <returns>
        /// Returns a string containing all the parameters formatted for an exec statement
        /// </returns>
        private string GetParameterDeclarationString()
        {
            const string Single_param = "{0} {1}, ";
            string parameterString = string.Empty;
            foreach (SqlParameter param in this.Parameters)
            {
                switch (param.DbType)
                {
                    case System.Data.DbType.String:
                    case System.Data.DbType.StringFixedLength:
                        parameterString += string.Format(Single_param, param.ParameterName, "nvarchar(max)");
                        break;
                    case System.Data.DbType.Time:
                    case System.Data.DbType.Date:
                    case System.Data.DbType.DateTime:
                    case System.Data.DbType.DateTime2:
                        parameterString += string.Format(Single_param, param.ParameterName, "datetime");
                        break;
                    default:
                        parameterString += string.Format(Single_param, param.ParameterName, "float");
                        break;
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
