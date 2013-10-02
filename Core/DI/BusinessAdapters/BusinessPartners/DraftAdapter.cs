// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DraftAdapter.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the DraftAdapter type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.BusinessPartners
{
    using System;
    using SAPbobsCOM;

    /// <summary>
    /// The Draft Adapter
    /// </summary>
    public class DraftAdapter : DocumentAdapter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DraftAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="draft">The draft.</param>
        public DraftAdapter(Company company, Documents draft) : base(company, draft)
        {
        }

        public DraftAdapter(Company company)
            : base(company)
        {
        }

        /// <summary>
        /// Gets the type of document this is an adapter for
        /// </summary>
        /// <value>The object type</value>
        public override BoObjectTypes DocumentType
        {
            get { return BoObjectTypes.oDrafts; }
        }

        /// <summary>
        /// Gets the target document type of this adapter's document type
        /// </summary>
        /// <value>The Target type</value>
        public override BoObjectTypes TargetType
        {
            get { throw new NotImplementedException(); }
        }
    }
}
