//-----------------------------------------------------------------------
// <copyright file="OrderAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.Sales
{
    #region Using Directive(s)

    using System;
    using System.Text;
    using SAPbobsCOM;

    #endregion Using Directive(s)

    public class QuoteAdapter : ARDocumentAdapter
    {
        public QuoteAdapter(Company company)
            : base(company)
        {            
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="QuoteAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="document">The document.</param>
        public QuoteAdapter(Company company, Documents document)
            : base(company, document)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="QuoteAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index.</param>
        public QuoteAdapter(Company company, string xmlFile, int index)
            : base(company, xmlFile, index)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="QuoteAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="docEntry">The doc entry.</param>
        public QuoteAdapter(Company company, int docEntry)
            : base(company, docEntry)
        {
        }


        /// <summary>
        /// Gets the type of document this is an adapter for
        /// </summary>
        /// <value></value>
        public override BoObjectTypes DocumentType
        {
            get { return BoObjectTypes.oQuotations; }
        }

        /// <summary>
        /// Gets the target document type of this adapter's document type
        /// </summary>
        /// <value></value>
        public override BoObjectTypes TargetType
        {
            get { return BoObjectTypes.oOrders; }
        }
    }
}
