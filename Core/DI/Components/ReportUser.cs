// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ReportUser.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   The report User
// </summary>
// --------------------------------------------------------------------------------------------------------------------
namespace B1C.SAP.DI.Components
{
    using Helpers;

    /// <summary>
    /// The report User
    /// </summary>
    public class ReportUser
    {
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>The user name.</value>
        public string Name
        {
            get
            {
                string reportUserName;
                ConfigurationHelper.Load("ReportUserName", out reportUserName);

                return reportUserName;
            }
        }

        /// <summary>
        /// Gets the password.
        /// </summary>
        /// <value>The user password.</value>
        public string Password
        {
            get
            {
                string reportUserPassword;
                ConfigurationHelper.Load("ReportUserPassword", out reportUserPassword);

                return reportUserPassword;
            }
        }
    }
}
