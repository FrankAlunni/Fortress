//-----------------------------------------------------------------------
// <copyright file="RequestedProductionOrder.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
//-----------------------------------------------------------------------

namespace Fortress.DI.Model.ProductionOrder
{

    #region Using Directive(s)

    using System.Collections.Generic;

    #endregion Using Directive(s)

    /// <summary>
    /// This is an abstract holding class for information regarding Production Orders
    /// that need to be created. This class can be used for delaying the creation of many 
    /// ProductionOrders until the full ProductionOrders hierarchy is set.
    /// </summary>
    public abstract class RequestedProductionOrder
    {
        /// <summary>
        /// Gets or sets the children requested production orders.
        /// </summary>
        /// <value>The children.</value>
        public IList<RequestedProductionOrder> Children
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets or sets the sales order id.
        /// </summary>
        /// <value>The sales order id.</value>
        public int SalesOrderId 
        { 
            get; 
            private set;
        }

        /// <summary>
        /// Gets or sets the sales order line.
        /// </summary>
        /// <value>The sales order line.</value>
        public int SalesOrderLine { get; private set; }

        /// <summary>
        /// Gets the drawing version.
        /// </summary>
        public int DrawingVersion { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="RequestedProductionOrder"/> class.
        /// </summary>
        protected RequestedProductionOrder()
        {
            Children = new List<RequestedProductionOrder>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RequestedProductionOrder"/> class.
        /// </summary>
        /// <param name="salesOrder">The sales order.</param>
        /// <param name="salesOrderLine">The sales order line.</param>
        protected RequestedProductionOrder(int salesOrder, int salesOrderLine, int drawingVersion)
        {
            this.SalesOrderId = salesOrder;
            this.SalesOrderLine = salesOrderLine;
            this.DrawingVersion = drawingVersion;
            Children = new List<RequestedProductionOrder>();
        }

        /// <summary>
        /// Adds a child RequestedProductionOrder.
        /// </summary>
        /// <param name="rpo">The child.</param>
        public void AddChild(RequestedProductionOrder rpo)
        {
            Children.Add(rpo);
        }

        /// <summary>
        /// Determines whether this instance has children.
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if this instance has children; otherwise, <c>false</c>.
        /// </returns>
        public bool HasChildren()
        {
            return Children.Count > 0;
        }

        /// <summary>
        /// Prints the Request.
        /// </summary>
        /// <param name="depth">The depth.</param>
        /// <returns>A string representing the request</returns>
        public abstract string Print(int depth);

        /// <summary>
        /// Returns a <see cref="System.String"/> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String"/> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return this.Print(0);
        }
    }
}