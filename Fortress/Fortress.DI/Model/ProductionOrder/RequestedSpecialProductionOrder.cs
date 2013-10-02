//-----------------------------------------------------------------------
// <copyright file="RequestedSpecialProductionOrder.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
//-----------------------------------------------------------------------

namespace Fortress.DI.Model.ProductionOrder
{
    #region Using Directive(s)
    
    using System.Text;

    #endregion Using Directive(s)

    public class RequestedSpecialProductionOrder : RequestedProductionOrder
    {

        /// <summary>
        /// Gets or sets the location of the XML file to use to create the Special Production Order.
        /// </summary>
        /// <value>The file location.</value>
        public string FileLocation
        {
            get;
            protected set;
        }

        /// <summary>
        /// Gets or sets the index of the XML.
        /// </summary>
        /// <value>The index of the XML.</value>
        public int XmlIndex
        {
            get;
            protected set;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RequestedSpecialProductionOrder"/> class.
        /// </summary>
        /// <param name="salesOrderId">The sales order id.</param>
        /// <param name="saleSorderLine">The sale sorder line.</param>
        /// <param name="fileLocation">The file location.</param>
        public RequestedSpecialProductionOrder(int salesOrderId, int saleSorderLine, string fileLocation, int drawingVersion)
            : base(salesOrderId, saleSorderLine, drawingVersion)
        {
            this.FileLocation = fileLocation;
            this.XmlIndex = 0;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RequestedSpecialProductionOrder"/> class.
        /// </summary>
        /// <param name="salesOrderId">The sales order id.</param>
        /// <param name="saleSorderLine">The sale sorder line.</param>
        /// <param name="fileLocation">The file location.</param>
        /// <param name="xmlIndex">Index of the XML.</param>
        public RequestedSpecialProductionOrder(int salesOrderId, int saleSorderLine, string fileLocation, int drawingVersion, int xmlIndex)
            : base(salesOrderId, saleSorderLine, drawingVersion)
        {
            this.FileLocation = fileLocation;
            this.XmlIndex = xmlIndex;
        }

        /// <summary>
        /// Prints the Request.
        /// </summary>
        /// <param name="depth">The depth.</param>
        /// <returns>A string representing the request</returns>
        public override string Print(int depth)
        {
            var builder = new StringBuilder();

            for (int i = 0; i < depth; i++)
            {
                builder.Append("    ");
            }

            builder.Append(this.FileLocation);
            builder.AppendLine();

            foreach (var child in this.Children)
            {
                builder.Append(child.Print(depth + 1));
            }

            return builder.ToString();
        }

    }
}
