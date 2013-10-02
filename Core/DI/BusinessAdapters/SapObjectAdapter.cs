//-----------------------------------------------------------------------
// <copyright file="SapObjectAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters
{
    #region Using Directive(s)

    using Helpers;
    using SAPbobsCOM;
    using Utility.Logging;    

    #endregion Using Dircetive(s)

    /// <summary>
    /// This is the abstract SAP Object Adapter
    /// </summary>
    public abstract class SapObjectAdapter : SapBaseObject
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="SapObjectAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        protected SapObjectAdapter(Company company)
        {
            this.Company = company;
            this.Exchanger = ExchangeHelper.Instance(this.Company);
        }
        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets or sets the task.
        /// </summary>
        /// <value>The running task.</value>
        public string Task
        {
            get
            {
                return ThreadedAppLog.GetCurrentTask();
            }

            set
            {
                ThreadedAppLog.SetTask(value);
            }
        }

        /// <summary>
        /// Gets or sets the exchanger
        /// </summary>
        internal ExchangeHelper Exchanger { get; set; }

        /// <summary>
        /// Gets the company.
        /// </summary>
        /// <value>The company.</value>
        protected Company Company { get; private set; }

        #endregion Properties

        #region Methods
        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected abstract override void Release();

        #endregion Methods
    }
}
