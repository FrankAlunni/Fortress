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

    public class RequestedAddBomProductionOrder : RequestedProductionOrder
    {
        /// <summary>
        /// Gets or sets the item code to create a standard production order with.
        /// </summary>
        /// <value>The item code.</value>
        public string ItemCode
        {
            get;
            private set;
        }

        public double Quantity
        {
            get;
            private set;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RequestedStandardProductionOrder"/> class.
        /// </summary>
        /// <param name="salesOrderId">The sales order id.</param>
        /// <param name="salesOrderLine">The sales order line.</param>
        /// <param name="itemCode">The item code.</param>
        public RequestedAddBomProductionOrder(int salesOrderId, int salesOrderLine, string itemCode, int drawingVersion, double quantity)
            : base(salesOrderId, salesOrderLine, drawingVersion)
        {
            this.ItemCode = itemCode;
            this.Quantity = quantity;
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

            builder.Append(this.ItemCode);
            builder.AppendLine();

            foreach (var child in this.Children)
            {
                builder.Append(child.Print(depth + 1));
            }

            return builder.ToString();
        }
    }
}
