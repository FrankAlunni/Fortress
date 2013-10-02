//-----------------------------------------------------------------------
// <copyright file="ApprovalStatus.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
//-----------------------------------------------------------------------
namespace Fortress.DI.Model.Enums
{
    public enum ApprovalStatus
    {
        None,
        DriveWorks,
        Approval,
        Yes,
        No
    }

    public static class ApprovalStatusHelper
    {
        public static string GetStringValue(ApprovalStatus pos)
        {
            string value = string.Empty;

            switch (pos)
            {
                case ApprovalStatus.None:
                    value = "";
                    break;
                case ApprovalStatus.DriveWorks:
                    value = "DriveWorks";
                    break;
                case ApprovalStatus.Approval:
                    value = "Approval";
                    break;
                case ApprovalStatus.Yes:
                    value = "Yes";
                    break;
                case ApprovalStatus.No:
                    value = "No";
                    break;
            }
            return value;
        }

        public static ApprovalStatus GetEnumFromString(string value)
        {
            ApprovalStatus toReturn = ApprovalStatus.None;

            try
            {
                switch (value)
                {
                    case "None":
                        toReturn = ApprovalStatus.None;
                        break;
                    case "DriveWorks":
                        toReturn = ApprovalStatus.DriveWorks;
                        break;
                    case "Approval":
                        toReturn = ApprovalStatus.Approval;
                        break;
                    case "Yes":
                        toReturn = ApprovalStatus.Yes;
                        break;
                    case "No":
                        toReturn = ApprovalStatus.No;
                        break;
                }
            }
            catch
            {
                toReturn = ApprovalStatus.None;
            }

            return toReturn;
        }
    }
}
