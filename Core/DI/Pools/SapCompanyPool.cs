// -----------------------------------------------------------------------
// <copyright file="SapCompanyPool.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <summary>
//    Product        : Messaging Service
//    Author Email   : bryan.atkinson@B1C.com
//    Created Date   : 08/26/2009
//    Prerequisites  : - SAP Business One v.2007a
//                     - SQL Server 2005
//                     - Micorsoft .NET Framework 3.5
// </summary>
//-----------------------------------------------------------------------

using B1C.Utility.Logging;
using B1C.Utility.Pools;

namespace B1C.SAP.DI.Pools
{
    #region Using Directives

    using System;
    using System.Configuration;
    using SAPbobsCOM;
 
    #endregion Using Directives
    /// <summary>
    /// This class represents a pool of SAP Company objects. It is used
    /// to share a number of *Connected* Company objects so that each class
    /// requiring a Company does not need to instantiate and Connect a new
    /// Company object, which are time-consuming operations.
    /// </summary>
    public class SapCompanyPool : BaseObjectPool<Company>
    {
        #region Private Members

        /// <summary>
        /// The maximum number of times to attempt to retreive a connected 
        /// company
        /// </summary>
        private const int MAX_COMPANY_RETRY = 100;

        /// <summary>
        /// The locking object
        /// </summary>
        private static readonly object _lock = new object();

        /// <summary>
        /// The singleton instance of the pool
        /// </summary>
        private static SapCompanyPool _instance;

        #endregion Private Members

        /// <summary>
        /// Prevents a default instance of the <see cref="SapCompanyPool"/> class from being created.
        /// </summary>
        private SapCompanyPool()
        {
            this.LoadSapCompanyConfiguration();
        }

        /// <summary>
        /// Gets the name of the pool.
        /// </summary>
        public override string PoolName
        {
            get { return "SAP Company Pool"; }
        }

        #region SAP Server connection properties

        /// <summary>
        /// Gets or sets the SAP server name
        /// </summary>
        private string ServerName { get; set; }

        /// <summary>
        /// Gets or sets the type of SAP Server.
        /// </summary>
        private int ServerType { get; set; }

        /// <summary>
        /// Gets or sets The password used to login to the server.
        /// </summary>
        private string ServerPassword { get; set; }

        /// <summary>
        /// Gets or sets The username used to login to the server.
        /// </summary>
        private string ServerUserName { get; set; }

        /// <summary>
        /// Gets or sets the company name to login to on the server.
        /// </summary>
        private string CompanyName { get; set; }

        /// <summary>
        /// Gets or sets the username used to login to the company.
        /// </summary>
        private string CompanyUserName { get; set; }

        /// <summary>
        /// Gets or sets the password used to login to the company.
        /// </summary>
        private string CompanyPassword { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether a trusted connection is used.
        /// </summary>
        private bool UseTrusted { get; set; }

        /// <summary>
        /// Gets or sets the language.
        /// </summary>
        private int Language { get; set; }

        #endregion SAP Server connection properties

        /// <summary>
        /// Gets the singleton instance of SapCompanyPool
        /// </summary>
        /// <returns>
        /// The singleton instance of the pool
        /// </returns>
        public static SapCompanyPool Instance()
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    // The double null-check is necessary in multi-threaded environments,
                    // as it is possible that two threads could pass through the first
                    // null check, and then both end up creating a new instance of the pool.
                    // The alternative is to lock before doign the null-check, but locking 
                    // is an expensive operation, and so should be minimized.
                    if (_instance == null)
                    {
                        _instance = new SapCompanyPool();
                    }
                }
            }

            return _instance;
        }

        /// <summary>
        /// Acquires a connected SAP Company object from the pool
        /// </summary>
        /// <returns>A connected SAP Company object</returns>
        public override Company Acquire()
        {
            ThreadedAppLog.WriteLine("Acquiring SAP Company Object [Pool Size: {0}]", base.PoolSize);

            // Retrieve the next company object from the base pool
            Company company = base.Acquire();            

            // If the company object is not connected, remove it from the 
            // queue and get the next one
            if (!company.Connected)
            {
                int companyCount = 1;

                do
                {
                    // Limit the number of retries to avoid infinite looping
                    if (companyCount > MAX_COMPANY_RETRY)
                    {
                        throw new PoolObjectInstantiationException(
                            string.Format(
                                "Could not retrieve a connected SAP Company object from pool. Tried {0} times.",
                                MAX_COMPANY_RETRY));
                    }

                    this.DisposeObject(company);
                    company = base.Acquire();

                    companyCount++;
                }
                while (!company.Connected);
            }

            return company;
        }

        /// <summary>
        /// Creates a new object for inclusion in the pool
        /// </summary>
        /// <returns>A newly instantiated and *connected* SAP Company object</returns>
        protected override Company CreateObject()
        {
            ThreadedAppLog.WriteLine("Creating new SAP Company Object");
            Company company = 
                new Company
                    {
                        DbServerType = (BoDataServerTypes) this.ServerType,
                        Server = this.ServerName,
                        DbPassword = this.ServerPassword,
                        DbUserName = this.ServerUserName,
                        UseTrusted = this.UseTrusted,
                        CompanyDB = this.CompanyName,
                        UserName = this.CompanyUserName,
                        Password = this.CompanyPassword,
                        language = (BoSuppLangs) this.Language
                    };

            if (company.Connect() != 0)
            {
                int errorCode;
                string errorMessage;
                company.GetLastError(out errorCode, out errorMessage);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                company = null;

                throw new PoolObjectInstantiationException(string.Format("Could not create a new Company object for pool {0}. SAP Error [Code: {1}] - [Message: {2}].", this.PoolName, errorCode, errorMessage));
            }

            return company;
        }

        /// <summary>
        /// Disposes an SAP company from the pool. The Company
        /// is disconnected.
        /// </summary>
        /// <param name="obj">
        /// The company being removed from the pool
        /// </param>
        public override void DisposeObject(Company obj)
        {
            // Dispose of the object from the base pool
            base.DisposeObject(obj);

            try
            {
                // Disconnect the Company from the DB
                obj.Disconnect();
            }
            catch (Exception ex)
            {
                ThreadedAppLog.WriteLine("Unable to disconnect company object in SapCompanyPool.");
                ThreadedAppLog.WriteErrorLine(ex);
            }
            finally
            {
                // Release the COM object from memory
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;

                // Force Garbage Collection
                GC.Collect();
            }
        }

        public override void Release(Company obj)
        {
            base.Release(obj);

            ThreadedAppLog.WriteLine("Releasing SAP Company Object [Pool Size: {0}]", base.PoolSize);
        }

        /// <summary>
        /// Loads the SAP company information
        /// </summary>
        private void LoadSapCompanyConfiguration()
        {
            this.ServerName = ConfigurationManager.AppSettings["DbServerName"];
            this.ServerType = Convert.ToInt32(ConfigurationManager.AppSettings["DbServerType"]);
            this.ServerPassword = ConfigurationManager.AppSettings["DbPassword"];
            this.ServerUserName = ConfigurationManager.AppSettings["DbUserName"];
            this.CompanyName = ConfigurationManager.AppSettings["Company"];
            this.CompanyUserName = ConfigurationManager.AppSettings["UserName"];
            this.CompanyPassword = ConfigurationManager.AppSettings["Password"];
            this.UseTrusted = Convert.ToBoolean(ConfigurationManager.AppSettings["UseTrusted"]);
            this.Language = Convert.ToInt32(ConfigurationManager.AppSettings["Language"]);
        }
    }
}