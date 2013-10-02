// --------------------------------------------------------------------------------------------------------------------
// <copyright file="UserAdapter.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the User type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.Administration
{
    using System;
    using Helpers;
    using SAPbobsCOM;

    /// <summary>
    /// The user information
    /// </summary>
    public class UserAdapter : SapObjectAdapter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UserAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="internalKey">The internal key.</param>
        public UserAdapter(Company company, int internalKey) : base(company)
        {

            using(var rs = new RecordsetAdapter(this.Company, string.Format("SELECT * FROM OUSR Where INTERNAL_K = {0}", internalKey)))
            {
                if (rs.EoF)
                {
                    this.InternalKey = 0;
                }
                else
                {
                    this.InternalKey = Convert.ToInt32(rs.FieldValue("INTERNAL_K").ToString());
                    this.Branch = Convert.ToInt32(rs.FieldValue("Branch").ToString());
                    this.Department = Convert.ToInt32(rs.FieldValue("Department").ToString());
                    this.EMail = rs.FieldValue("E_Mail").ToString();
                    this.FaxNumber = rs.FieldValue("Fax").ToString();
                    this.Locked = rs.FieldValue("Locked").ToString() == "Y";
                    this.Superuser = rs.FieldValue("SUPERUSER").ToString() == "Y";
                    this.UserCode = rs.FieldValue("U_NAME").ToString();
                    this.UserName = rs.FieldValue("USER_CODE").ToString();
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UserAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="userCode">The user code.</param>
        public UserAdapter(Company company, string userCode) : base(company)
        {
            using (var rs = new RecordsetAdapter(this.Company, string.Format("SELECT * FROM OUSR Where USER_CODE = '{0}'", userCode)))
            {
                if (rs.EoF)
                {
                    this.InternalKey = 0;
                }
                else
                {
                    this.InternalKey = Convert.ToInt32(rs.FieldValue("INTERNAL_K").ToString());
                    this.Branch = Convert.ToInt32(rs.FieldValue("Branch").ToString());
                    this.Department = Convert.ToInt32(rs.FieldValue("Department").ToString());
                    this.EMail = rs.FieldValue("E_Mail").ToString();
                    this.FaxNumber = rs.FieldValue("Fax").ToString();
                    this.Locked = rs.FieldValue("Locked").ToString() == "Y";
                    this.Superuser = rs.FieldValue("SUPERUSER").ToString() == "Y";
                    this.UserCode = rs.FieldValue("U_NAME").ToString();
                    this.UserName = rs.FieldValue("USER_CODE").ToString();
                }

            }
        }

        /// <summary>
        /// Gets or sets the branch.
        /// </summary>
        /// <value>The branch.</value>
        public int Branch { get; set; }

        /// <summary>
        /// Gets or sets the department.
        /// </summary>
        /// <value>The department.</value>
        public int Department { get; set; }

        /// <summary>
        /// Gets or sets the Email.
        /// </summary>
        /// <value>The Email.</value>
        public string EMail { get; set; }

        /// <summary>
        /// Gets or sets the fax number.
        /// </summary>
        /// <value>The fax number.</value>
        public string FaxNumber { get; set; }

        /// <summary>
        /// Gets or sets the internal key.
        /// </summary>
        /// <value>The internal key.</value>
        public int InternalKey { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="UserAdapter"/> is locked.
        /// </summary>
        /// <value><c>true</c> if locked; otherwise, <c>false</c>.</value>
        public bool Locked { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="UserAdapter"/> is superuser.
        /// </summary>
        /// <value><c>true</c> if superuser; otherwise, <c>false</c>.</value>
        public bool Superuser { get; set; }

        /// <summary>
        /// Gets or sets the user code.
        /// </summary>
        /// <value>The user code.</value>
        public string UserCode { get; set; }

        /// <summary>
        /// Gets or sets the name of the user.
        /// </summary>
        /// <value>The name of the user.</value>
        public string UserName { get; set; }

        /// <summary>
        /// Currents the specified company.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns>The User adapter for the logged user</returns>
        public static UserAdapter Current(Company company)
        {
            var user = new UserAdapter(company, company.UserSignature);
            if (user.IsLoaded())
            {
                return user;
            }

            return null;
        }

        /// <summary>
        /// Determines whether this instance is loaded.
        /// </summary>
        /// <returns>
        /// <c>true</c> if this instance is loaded; otherwise, <c>false</c>.
        /// </returns>
        public bool IsLoaded()
        {
            return this.InternalKey > 0;
        }

        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected override void Release()
        {
            return;
        }
    }
}
