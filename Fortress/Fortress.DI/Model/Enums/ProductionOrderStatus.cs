//-----------------------------------------------------------------------
// <copyright file="ProductionOrderStatus.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
//-----------------------------------------------------------------------

namespace Fortress.DI.Model.Enums
{
    public enum ProductionOrderStatus
    {
        None,
        SpecialProductionOrderSuccess,
        SpecialProductionOrderFailed,
        StandardProductionOrderSuccess,
        StandardProductionOrderFailed
    }

    public static class ProductionOrderStatusHelper
    {
        public static string GetStringValue(ProductionOrderStatus pos)
        {
            string value = string.Empty;

            switch (pos)
            {
                case ProductionOrderStatus.None:
                    value = "";
                    break;
                case ProductionOrderStatus.SpecialProductionOrderSuccess:
                    value = "SpSuccess";
                    break;
                case ProductionOrderStatus.SpecialProductionOrderFailed:
                    value = "SpFailed";
                    break;
                case ProductionOrderStatus.StandardProductionOrderFailed:
                    value = "BomFailed";
                    break;
                case ProductionOrderStatus.StandardProductionOrderSuccess:
                    value = "BomSuccess";
                    break;
            }
            return value;
        }

        public static ProductionOrderStatus GetEnumFromString(string value)
        {
            ProductionOrderStatus toReturn = ProductionOrderStatus.None;

            try
            {
                switch (value)
                {
                    case "SpSuccess":
                        toReturn = ProductionOrderStatus.SpecialProductionOrderSuccess;
                        break;
                    case "SpFailed":
                        toReturn = ProductionOrderStatus.SpecialProductionOrderFailed;
                        break;
                    case "BomSuccess":
                        toReturn = ProductionOrderStatus.StandardProductionOrderSuccess;
                        break;
                    case "BomFailed":
                        toReturn = ProductionOrderStatus.StandardProductionOrderFailed;
                        break;
                }
            }
            catch
            {
                toReturn = ProductionOrderStatus.None;
            }

            return toReturn;
        }
    }
}
