//-----------------------------------------------------------------------
// <copyright file="DatabaseUpdater.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
// <email>frank.alunni@b1computing.com</email>
//-----------------------------------------------------------------------

namespace B1C.SAP.Addons.Configurator.Model
{
    #region Using Directive(s)

    using System;
    using System.Collections.Generic;
    using SAPbobsCOM;
    using DI.BusinessAdapters;
    using DI.Business;

    #endregion Using Directive(s)

    public static class DatabaseUpdater
    {
        /// <summary>
        /// Creates the general settings.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <returns>true, if successful</returns>
        public static bool CreateCommission(AddOn addOn)
        {
            try
            {
                var table = new TableAdapter(addOn.Company, "XX_COMMISS") { Description = "Commission" };


                table.Type = BoUTBTableType.bott_MasterData;

                table.Fields.Add(addOn.Company, "XX_COMMISS", "XX_OrderNo", "Sales Order Number", 0, 0, BoFieldTypes.db_Numeric, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_COMMISS", "XX_OrdrLnNo", "Sales Order Line Number", 0, 0, BoFieldTypes.db_Numeric, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_COMMISS", "XX_CommisAv", "Total Commission Avail.", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Rate, string.Empty);

                var validValues = new Dictionary<string, string>();
                validValues.Add("Full", "Full");
                validValues.Add("Engineering", "Engineering");
                validValues.Add("Origin", "Origin");
                validValues.Add("Destination", "Destination");
                validValues.Add("Service", "Service");
                table.Fields.Add(addOn.Company, "XX_COMMISS", "XX_Type", "Type", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_COMMISS", "XX_Rate", "Rate", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Rate, string.Empty);

                validValues = new Dictionary<string, string>();
                validValues.Add("None", "None");
                validValues.Add("Approved", "Approved");
                table.Fields.Add(addOn.Company, "XX_COMMISS", "XX_Approved", "Approved", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                table.Registration.ObjectName = "Commission";
                table.Registration.ObjectCode = "Commission";
                table.Registration.CanCancel = true;
                table.Registration.CanClose = true;
                table.Registration.CanDelete = true;
                table.Registration.CanFind = true;
                table.Registration.CanLog = true;
                table.Registration.FindColumns.Add("Name");
                table.Registration.FindColumns.Add("Code");

                table.StatusbarChanged += addOn.StatusbarChanged;
                table.Add();
                table.StatusbarChanged -= addOn.StatusbarChanged;

                return true;
            }
            catch (Exception ex)
            {
                addOn.HandleException(ex);
                return false;
            }
        }

        /// <summary>
        /// Creates the general settings.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <returns>true if successful</returns>
        public static bool CreateApproval(AddOn addOn)
        {
            try
            {
                var table = new TableAdapter(addOn.Company, "XX_APPROV") { Description = "Approval" };


                table.Type = BoUTBTableType.bott_MasterData;

                table.Fields.Add(addOn.Company, "XX_APPROV", "XX_OrderNo", "Sales Order Number", 0, 0, BoFieldTypes.db_Numeric, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_APPROV", "XX_OrdrLnNo", "Sales Order Line Number", 0, 0, BoFieldTypes.db_Numeric, null, string.Empty);

                var validValues = new Dictionary<string, string>();
                validValues.Add("-", "-");
                validValues.Add("Drawing Required", "Drawing Required");
                validValues.Add("App Drawing Completed", "App Drawing Completed");
                validValues.Add("Sent to Customer", "Sent to Customer");
                validValues.Add("Changes Required", "Changes Required");
                validValues.Add("Approved", "Approved");
                validValues.Add("Prod. Drawing Required", "Prod. Drawing Required");
                validValues.Add("Prod. Drawing Completed", "Prod. Drawing Completed");
                validValues.Add("On Hold", "On Hold");
                table.Fields.Add(addOn.Company, "XX_APPROV", "XX_Status", "Status", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);


                table.Registration.ObjectName = "Approval Process";
                table.Registration.ObjectCode = "Approval";
                table.Registration.CanCancel = true;
                table.Registration.CanClose = true;
                table.Registration.CanDelete = true;
                table.Registration.CanFind = true;
                table.Registration.CanLog = true;
                table.Registration.FindColumns.Add("Name");
                table.Registration.FindColumns.Add("Code");

                table.StatusbarChanged += addOn.StatusbarChanged;
                table.Add();
                table.StatusbarChanged -= addOn.StatusbarChanged;

                return true;
            }
            catch (Exception ex)
            {
                addOn.HandleException(ex);
                return false;
            }
        }

        /// <summary>
        /// Creates the general settings.
        /// </summary>
        /// <returns>true if successful</returns>
        public static bool CreateGeneralSettings(AddOn addOn)
        {
            try
            {
                var table = new TableAdapter(addOn.Company, "XX_GENSET") { Description = "General Settings" };

                table.Load();

                if (table.Exists)
                {
                    return false;
                }

                table.Type = BoUTBTableType.bott_MasterData;

                table.Fields.Add(addOn.Company, "XX_GENSET", "XX_Value", "Value", 254, 254, BoFieldTypes.db_Alpha, null, string.Empty);

                table.Registration.ObjectName = "General Settings";
                table.Registration.ObjectCode = "GeneralSettings";
                table.Registration.CanCancel = true;
                table.Registration.CanClose = true;
                table.Registration.CanDelete = true;
                table.Registration.CanFind = true;
                table.Registration.CanLog = true;
                table.Registration.FindColumns.Add("Code");
                table.Registration.FindColumns.Add("Name");
                table.Registration.FindColumns.Add("U_XX_Value");
                table.Registration.CanCreateDefaultForm = true;
                table.Registration.FormColumns.Add("Code");
                table.Registration.FormColumns.Add("Name");
                table.Registration.FormColumns.Add("U_XX_Value");

                table.StatusbarChanged += addOn.StatusbarChanged;
                table.Add();
                table.StatusbarChanged -= addOn.StatusbarChanged;


                // Populate Data
                var loadData = new System.Text.StringBuilder();
                loadData.AppendLine("<DataLoader UDOName='GeneralSettings' UDOType='MasterData'>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "DetPref");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Metal detector item code prefix");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "D-");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "ConvPref");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Conveyor belt item code prefix");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "C-");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "SODirPtn");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Sales Order directory name pattern");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "SO {DocNumber:#00000}");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "LnDirPtn");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Line directory name pattern");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "{LineNum:#00}[{ItemCode}]");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "SpBomPtn");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Special BOM Naming Convention");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "SO{DocNumber:#00000}-{LineNum:#00}-{Revision:#000} [{ItemCode}].xml");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "SpBomDir");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Special BOM(s) Path");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "\\\\bastion\\B1C Product Configurator\\XML Files\\");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "ConfNtf");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Configuration change notification user");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "[Enter User to be nofified]");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "ConfSubj");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Configuration change subject");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "SO{DocNumber} - {LineNum} - {ItemCode} revised");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "ConfBody");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Configuration change body");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "The configuration for Sales Order {DocNumber}, line {LineNum} - {ItemCode} has been revised from revision {oldRev} to {newRev}");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "AutoSel");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Flag to auto-select first row");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "true");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "SoDirSt");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Drawing Dir Starting Range");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "10000");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "SoDirFa");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Drawing Dir Div Factor");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "10");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "SoDirFrm");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Drawing Dir Number Format");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "00000");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "MdComRt");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Default Metal Detector Commision Pct.");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "15");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("<Record>");
                loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", "CnComRt");
                loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", "Default Conveyor Belt Commision Pct.");
                loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", "10");
                loadData.AppendLine("</Record>");
                loadData.AppendLine("</DataLoader>");

                var loader = DataLoader.InstanceFromXMLString(loadData.ToString());
                if (loader != null)
                {
                    loader.Process(addOn.Company);

                    if (!string.IsNullOrEmpty(loader.ErrorMessage))
                    {
                        throw new Exception("Data loader exception: " + loader.ErrorMessage);
                    }
                }


                return true;
            }
            catch (Exception ex)
            {
                addOn.HandleException(ex);
                return false;
            }
        }

        /// <summary>
        /// Creates the serial numbers.
        /// </summary>
        /// <returns>true if successful</returns>
        public static bool CreateMetalDetectorConfigurator(AddOn addOn)
        {
            try
            {
                var table = new TableAdapter(addOn.Company, "XX_METDET") { Description = "Metal Detector Configuration" };

                if (table.Exists)
                {
                    return false;
                }

                table.Type = BoUTBTableType.bott_MasterData;

                var distanceUnits = new Dictionary<string, string>();
                distanceUnits.Add("ft", "Feet");
                distanceUnits.Add("in", "Inches");
                distanceUnits.Add("m", "Meters");
                distanceUnits.Add("mm", "Millimiters");

                var yesNoList = new Dictionary<string, string>();
                yesNoList.Add("Y", "Y");
                yesNoList.Add("N", "N");

                // **************************************
                // General Information
                // **************************************
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_OrderNo", "Sales Order Number", 0, 0, BoFieldTypes.db_Numeric, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_OrdrLnNo", "Sales Order Line Number", 0, 0, BoFieldTypes.db_Numeric, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Revis", "Product Revision", 254, 254, BoFieldTypes.db_Alpha, null, string.Empty);

                var validValues = new Dictionary<string, string>();
                validValues.Add("Initiated", "Initiated");
                validValues.Add("Modified", "Modified");
                validValues.Add("Failed", "Failed");
                validValues.Add("Passed", "Passed");
                validValues.Add("Ready", "Ready");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Status", "Status", 25, 25, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                // Product Information Tab
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Descr", "Product Description", 254, 254, BoFieldTypes.db_Alpha, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_PkgMat", "Packaging Material", 254, 254, BoFieldTypes.db_Alpha, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_PrdPerMn", "Product Per Minute", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);

                validValues = new Dictionary<string, string>();
                validValues.Add("Conveyor", "Conveyor");
                validValues.Add("Gravity", "Gravity");
                validValues.Add("Pipeline", "Pipeline");
                validValues.Add("Pharmaceutical", "Pharmaceutical");
                validValues.Add("Surface", "Surface");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Applic", "Application", 25, 25, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                validValues = new Dictionary<string, string>();
                validValues.Add("Rect", "Rect");
                validValues.Add("Round", "Round");
                validValues.Add("Bulk", "Bulk");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_PkgSh", "Package Shape", 25, 25, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                validValues = new Dictionary<string, string>();
                validValues.Add("in^3", "in^3");
                validValues.Add("m^3", "Round");
                validValues.Add("lb", "lb");
                validValues.Add("kg", "kg");
                validValues.Add("gal", "gal");
                validValues.Add("Tablets", "Tablets");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_BlkFlRt", "Bulk Flow Rate Per Hour", 25, 25, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                validValues = new Dictionary<string, string>();
                validValues.Add("C°", "C°");
                validValues.Add("F°", "F°");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_TempU", "Temperature Unit", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Temp", "Temperature", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdWdMn", "Along Belt Width Min", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdWdMx", "Along Belt Width Max", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdHgMn", "Height Min", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdHgMx", "Height Max", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdLnMn", "Along Belt Length Min", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdLnMx", "Along Belt Length Max", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdDmMn", "Diameter Min", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdDmMx", "Diameter Max", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdWdMnU", "Along Belt Width Min Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdWdMxU", "Along Belt Width Max Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdHgMnU", "Height Min Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdHgMxU", "Height Max Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdLnMnU", "Along Belt Length Min Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdLnMxU", "Along Belt Length Max Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdDmMnU", "Diameter Min Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdDmMxU", "Diameter Max Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");

                validValues = new Dictionary<string, string>();
                validValues.Add("Lb", "Lb");
                validValues.Add("Oz", "Oz");
                validValues.Add("Kg", "Kg");
                validValues.Add("g", "g");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdWgUn", "Weight Unit", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdWgMx", "Product Weight Max", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ProdWgMn", "Product Weight Min", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);

                // Surface Unit Tab
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_WebWd", "Web Width", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_TrackEr", "Tracking Error", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Thick", "Thickness", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_FeedRt", "Feed Rate (ft/min)", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty);


                // Gravity Unit Tab
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Densit", "Density", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty);

                // Pharmaceutical Unit Tab
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_TabletWd", "Tablet Width", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_TabletWdU", "Tablet Width Unit", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_TabletHg", "Tablet Height", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_TabletHgU", "Tablet Height Unit", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");

                validValues = new Dictionary<string, string>();
                validValues.Add("Tablet", "Tablet");
                validValues.Add("Capsule", "Capsule");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_TabletSh", "Tablet Shape", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_IronEnr", "Iron Enriched", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");

                // Pipeline Unit Tab
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Viscos", "Viscosity", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Press", "Pressure", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty);

                //*****************************************
                // Metal Detector
                //*****************************************

                // Metal Detector Master Tab
                validValues = new Dictionary<string, string>();
                validValues.Add("Phantom", "Phantom");
                validValues.Add("Pinpoint", "Pinpoint");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Model", "Model", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_PartNo", "Part Number", 100, 100, BoFieldTypes.db_Alpha, null, string.Empty);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_BeltWd", "Belt Width", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_BeltWdU", "Belt Width Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");

                validValues = new Dictionary<string, string>();
                validValues.Add("Inch", "Inch");
                validValues.Add("Millimeter", "Millimeter");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_MDUnit", "Unit of Measure", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_BSH", "Brick Shit House", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");

                validValues = new Dictionary<string, string>();
                validValues.Add("None", "None");
                validValues.Add("Standard", "Standard");
                validValues.Add("SS Membrane", "SS Membrane");
                validValues.Add("Touch Screen", "Touch Screen");
                validValues.Add("Plastic Membrane", "Plastic Membrane");
                validValues.Add("Capacitive", "Capacitive");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_CtrlTyp", "Control Type", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, "Standard");

                validValues = new Dictionary<string, string>();
                validValues.Add("Rectangular", "Rectangular");
                validValues.Add("Circular", "Circular");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Style", "Style", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ApertWd", "Aperture Width", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ApertWdU", "Aperture Width Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ApertHg", "Aperture Height", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ApertHgU", "Aperture Height Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ApertDm", "Aperture Diameter", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ApertDmU", "Aperture Diameter Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ApertTC", "Aperture Thru Calc", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ApertBC", "Aperture Border Calc", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Freq1", "Detector Frequency 1", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_DualFreq", "Dual Frequency", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Freq2", "Detector Frequency 2", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty);

                validValues = new Dictionary<string, string>();
                validValues.Add("Standard", "Standard");
                validValues.Add("Gravity", "Gravity");
                validValues.Add("Pipeline", "Pipeline");
                validValues.Add("Surface", "Surface");
                validValues.Add("Pharmaceutical", "Pharmaceutical");
                validValues.Add("Vertex", "Vertex");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_DetType", "Detector Type", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                validValues = new Dictionary<string, string>();
                validValues.Add("Steel It", "Steel It");
                validValues.Add("White", "White");
                validValues.Add("SS", "SS");
                validValues.Add("RAL", "RAL");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_DetFnsh", "Detector Finish", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                validValues = new Dictionary<string, string>();
                validValues.Add("Dry", "Dry");
                validValues.Add("Wet", "Wet");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_PhMode", "Phase Mode", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                validValues = new Dictionary<string, string>();
                validValues.Add("Dry", "Dry");
                validValues.Add("Wet", "Wet");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Enviro", "Enivironment", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                validValues = new Dictionary<string, string>();
                validValues.Add("English", "English");
                validValues.Add("French", "French");
                validValues.Add("Spanish", "Spanish");
                validValues.Add("Portuguese", "Portuguese");
                validValues.Add("German", "German");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Lang", "Language", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                validValues = new Dictionary<string, string>();
                validValues.Add("Horizontal", "Horizontal");
                validValues.Add("Vertical", "Vertical");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_DetFlow", "Detector Flow", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                validValues = new Dictionary<string, string>();
                validValues.Add("Epoxy", "Epoxy");
                validValues.Add("Phenolic", "Phenolic");
                validValues.Add("Belting", "Belting");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_DetChute", "Detector Chute", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, "Phenolic");

                /*
                validValues = new Dictionary<string, string>();
                validValues.Add("FE", "FE");
                validValues.Add("NF", "NF");
                validValues.Add("SS", "SS");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Sensit", "Expected Sensitivity", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                */

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_SensitFE", "Expected FE Sensitivity Value", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_SensitNF", "Expected NF Sensitivity Value", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_SensitSS", "Expected SS Sensitivity Value", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);


                // Metal Detector Surface Tab
                validValues = new Dictionary<string, string>();
                validValues.Add("Up", "Up");
                validValues.Add("Down", "Down");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_SensDir", "Sensitivity Direction", 5, 5, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Shield", "Shielded", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, null);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_MultiSc", "Number of Sections", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_AlarmBR", "Alarms - Basic with Reset", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_AlarmMBO", "Alarms - Multiscan Built On", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_AlarmRR", "Alarms - Remote Reset", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");

                // Metal Detector Option Tab
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_RgRequ", "Ranges Required", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_RgSize", "Ranges Size", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_RgSizeU", "Ranges Size Units", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_RgContct", "Ranges Contact", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");

                validValues = new Dictionary<string, string>();
                validValues.Add("Ethernet", "Ethernet");
                validValues.Add("Wireless", "Wireless");
                validValues.Add("Serial", "Serial");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_RgContTp", "Ranges Contact Type", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_ExplPrf", "Explosion Proof", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_Remote", "Detector Remote", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");

                validValues = new Dictionary<string, string>();
                validValues.Add("Standard", "Standard");
                validValues.Add("SS Membrane", "SS Membrane");
                validValues.Add("Touch Screen", "Touch Screen");
                validValues.Add("Plastic Membrane", "Plastic Membrane");
                validValues.Add("Capacitive", "Capacitive");
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_RemoteTp", "Detector Remote Type", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                table.Fields.Add(addOn.Company, "XX_METDET", "XX_UNotes", "User Notes", 0, 0, BoFieldTypes.db_Memo, null, string.Empty, null, null);
                table.Fields.Add(addOn.Company, "XX_METDET", "XX_SNotes", "User Notes", 0, 0, BoFieldTypes.db_Memo, null, string.Empty, null, null);

                // Provide Registration Informations
                table.Registration.ObjectName = "Metal Detector Configurator";
                table.Registration.ObjectCode = "MDConfigurator";
                table.Registration.CanCancel = true;
                table.Registration.CanClose = true;
                table.Registration.CanDelete = true;
                table.Registration.CanFind = true;
                table.Registration.CanLog = true;
                table.Registration.FindColumns.Add("Name");
                table.Registration.FindColumns.Add("Code");

                table.StatusbarChanged += addOn.StatusbarChanged;
                table.Add();
                table.StatusbarChanged -= addOn.StatusbarChanged;

                return true;

            }
            catch (Exception ex)
            {
                addOn.HandleException(ex);
                return false;
            }
        }

        /// <summary>
        /// Creates the conveyor belt configurator.
        /// </summary>
        /// <returns>true if successful</returns>
        public static bool CreateConveyorBeltConfigurator(AddOn addOn)
        {
            try
            {
                var table = new TableAdapter(addOn.Company, "XX_CONVBLT") { Description = "Conveyor Belt Configuration" };

                if (table.Exists)
                {
                    return false;
                }

                table.Type = BoUTBTableType.bott_MasterData;

                var distanceUnits = new Dictionary<string, string>();
                distanceUnits.Add("ft", "Feet");
                distanceUnits.Add("in", "Inches");
                distanceUnits.Add("m", "Meters");
                distanceUnits.Add("mm", "Millimiters");

                var yesNoList = new Dictionary<string, string>();
                yesNoList.Add("Y", "Y");
                yesNoList.Add("N", "N");

                // **************************************
                // General Information
                // **************************************
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_OrderNo", "Sales Order Number", 0, 0, BoFieldTypes.db_Numeric, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_OrdrLnNo", "Sales Order Line Number", 0, 0, BoFieldTypes.db_Numeric, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_Revis", "Product Revision", 254, 254, BoFieldTypes.db_Alpha, null, string.Empty);

                var validValues = new Dictionary<string, string>();
                validValues.Add("Initiated", "Initiated");
                validValues.Add("Modified", "Modified");
                validValues.Add("Failed", "Failed");
                validValues.Add("Passed", "Passed");
                validValues.Add("Ready", "Ready");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_Status", "Status", 25, 25, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                /*********************************
                 * Conveyor Master Tab
                 * *******************************/
                // Belt
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_OvLenght", "Overall Length", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                validValues = new Dictionary<string, string>();
                validValues.Add("Plastic Chain Open Grid", "Plastic Chain Open Grid");
                validValues.Add("Plastic Chain Flat Top", "Plastic Chain Flat Top");
                validValues.Add("Intralox", "Intralox");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BltType", "Belt Type", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BltSpeed", "Belt Speed", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BltWd", "Belt Width", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BltPM", "Belt Product/Min", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BltLn", "Belt Line", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                validValues = new Dictionary<string, string>();
                validValues.Add("Right to Left", "Right to Left");
                validValues.Add("Left to Right", "Left to Right");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BltDir", "Belt Direction of Travel", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, "Left to Right");
                // Belt Height
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BltHgIn", "Belt Height In", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BltHgOut", "Belt Height Out", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                validValues = new Dictionary<string, string>();
                validValues.Add("2", "2");
                validValues.Add("4", "4");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BltAdj", "Adjustment Range", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                // Detector Location
                validValues = new Dictionary<string, string>();
                validValues.Add("Left", "Left");
                validValues.Add("Right", "Right");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_DetLoc", "Detector Location", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_DetLocC", "Detector Location Centered", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_DetFrCtr", "Detector Location From Center", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty);
                // Mounting Info
                validValues = new Dictionary<string, string>();
                validValues.Add("Fixed Feet", "Fixed Feet");
                validValues.Add("Casters", "Casters");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_MntType", "Mount Type", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_MntStyle", "Swivel", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                validValues = new Dictionary<string, string>();
                validValues.Add("Zinc", "Zinc");
                validValues.Add("SS", "SS");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_MntMat", "Mount Material", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                // Frame - Out Riggers
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_OutRReq", "Outriggers Required", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                validValues = new Dictionary<string, string>();
                validValues.Add("Spidered", "Spidered");
                validValues.Add("Angular", "Angular");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_OutRStl", "Outriggers Style", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, "Angular");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_OutLn", "Outriggers Length", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_OutLnU", "Outriggers Length Unit", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_OutWd", "Outriggers Width", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_OutWdU", "Outriggers Width Unit", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_OutHg", "Outriggers Height", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, string.Empty);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_OutHgU", "Outriggers Height Unit", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, distanceUnits, "in");
                validValues = new Dictionary<string, string>();
                validValues.Add("STD", "STD");
                validValues.Add("BBK", "BBK");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_FrmType", "Frame Type", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("Standard", "Standard");
                validValues.Add("Infeed", "Infeed");
                validValues.Add("Discharge", "Discharge");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_CtrlLoc", "Control Location", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                // Motor/Gear
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_MtrReq", "Motor Required", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_SngDrop", "Single Drop", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                validValues = new Dictionary<string, string>();
                validValues.Add("110", "110");
                validValues.Add("220", "220");
                validValues.Add("460", "460");
                validValues.Add("575", "575");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_Voltage", "Supplied Voltage", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("3 Phase In", "3 Phase In");
                validValues.Add("1 Phase In", "1 Phase In");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_Phase", "Phase", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, "1 Phase In");
                validValues = new Dictionary<string, string>();
                validValues.Add("0.33 HP", "0.33 HP");
                validValues.Add("0.5 HP", "0.5 HP");
                validValues.Add("0.75 HP", "0.75 HP");
                validValues.Add("1 HP", "1 HP");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_HP", "Horse Power", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("208-460V, 3 Phase, 1/3HP", "208-460V, 3 Phase, 1/3HP");
                validValues.Add("208-460V, 3 Phase, 1/2HP", "208-460V, 3 Phase, 1/2HP");
                validValues.Add("208-460V, 3 Phase, 1/2HP, SS", "208-460V, 3 Phase, 1/2HP, SS");
                validValues.Add("575V, 3 Phase, 1/2HP", "575V, 3 Phase, 1/2HP");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_MotType", "Motor Type", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_HrzMot", "Horzontal Motor", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_SlvDrive", "Slave Drive", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                validValues = new Dictionary<string, string>();
                validValues.Add("KELTECH GEARBOX SS 56C 14:1", "KELTECH GEARBOX SS 56C 14:1");
                validValues.Add("KELTECH GEARBOX SS 56C 36:1", "KELTECH GEARBOX SS 56C 36:1");
                validValues.Add("KELTECH GEARBOX SS 56C 43:1", "KELTECH GEARBOX SS 56C 43:1");
                validValues.Add("NORD GEARBOX 56C 10:1", "NORD GEARBOX 56C 10:1");
                validValues.Add("NORD GEARBOX 56C 15:1", "NORD GEARBOX 56C 15:1");
                validValues.Add("NORD GEARBOX 56C 20:1", "NORD GEARBOX 56C 20:1");
                validValues.Add("NORD GEARBOX 56C 25:1", "NORD GEARBOX 56C 25:1");
                validValues.Add("NORD GEARBOX 56C 30:1", "NORD GEARBOX 56C 30:1");
                validValues.Add("NORD GEARBOX 56C 40:1", "NORD GEARBOX 56C 40:1");
                validValues.Add("NORD GEARBOX 56C 5:1", "NORD GEARBOX 56C 5:1");
                validValues.Add("NORD GEARBOX 56C 50:1", "NORD GEARBOX 56C 50:1");
                validValues.Add("NORD GEARBOX 56C 60:1", "NORD GEARBOX 56C 60:1");
                validValues.Add("NORD GEARBOX 56C 7.5:1", "NORD GEARBOX 56C 7.5:1");
                validValues.Add("NORD GEARBOX 56C 80:1", "NORD GEARBOX 56C 80:1");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_GearBD", "Gear Box Drive", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                // Options
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RejReq", "Reject Required", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("Inch", "Inch");
                validValues.Add("Millimeter", "Millimeter");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_Unit", "Unit of Measure", 15, 15, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_EmergStop", "Emergency Stop", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_InFeedEy", "In Feed Eye", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_EncReq", "Encoder Required", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                validValues = new Dictionary<string, string>();
                validValues.Add("None", "None");
                validValues.Add("Beads", "Beads");
                validValues.Add("Dead Plate", "Dead Plate");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_TraBead", "Transfer Beads", 15, 15, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("Both", "Both");
                validValues.Add("Infeed", "Infeed");
                validValues.Add("Discharge", "Discharge");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_TraSide", "Transfer Side", 15, 15, BoFieldTypes.db_Alpha, null, string.Empty, validValues, "Infeed");
                validValues = new Dictionary<string, string>();
                validValues.Add("None", "None");
                validValues.Add("Infeed", "Infeed");
                validValues.Add("Discharge", "Discharge");
                validValues.Add("Both", "Both");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BmpGrd", "Bumper Guard", 15, 15, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("None", "None");
                validValues.Add("Edwards", "Edwards");
                validValues.Add("Edwards Class 2 Div 2", "Edwards Class 2 Div 2");
                validValues.Add("Allen Bradley", "Allen Bradley");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_Alarm", "Alarm", 25, 25, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                // Guide Rails
                validValues = new Dictionary<string, string>();
                validValues.Add("None", "None");
                validValues.Add("Arboron", "Arboron");
                validValues.Add("Skirted", "Skirted");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RailType", "Guide Rails Type", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("Infeed", "Infeed");
                validValues.Add("Full", "Full");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RailLn", "Guide Rails Length", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("2", "2");
                validValues.Add("4", "4");
                validValues.Add("Other", "Other");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RailHg", "Guide Rails Height", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RailHgOpt", "Guide Rails Height Options", 100, 100, BoFieldTypes.db_Alpha, null, string.Empty);

                // Starter
                validValues = new Dictionary<string, string>();
                validValues.Add("None", "None");
                validValues.Add("VFD", "VFD");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_StType", "Starter Type", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_StOptPt", "Starter Option Potentiometer", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_StOptKR", "Starter Option Key Reset", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_StOptBR", "Starter Option Belt Reverse", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_StOptBS", "Starter Option Belt Switches", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_StOptES", "Starter Option EStop", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("Soft", "Soft");
                validValues.Add("Relay", "Relay");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_LatcType", "Latch Type", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                // Light Stack
                validValues = new Dictionary<string, string>();
                validValues.Add("None", "None");
                validValues.Add("EZLite", "EZLite");
                validValues.Add("Stack", "Stack");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_LSType", "Light Stack Type", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("Infeed CS", "Infeed CS");
                validValues.Add("Infeed OC", "Infeed OC");
                validValues.Add("Discharge CS", "Discharge CS");
                validValues.Add("Discharge OC", "Discharge OC");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_LSLoc", "Light Stack Location", 20, 20, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_LSRed", "Light Stack Red", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_LSAmber", "Light Stack Amber", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_LSGreen", "Light Stack Green", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_LSLtc", "Light Stack Latching", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_LSBuzzer", "Light Stack Buzzer", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");

                /*********************************
                 * Reject Tab
                 * *******************************/
                validValues = new Dictionary<string, string>();
                validValues.Add("Belt Stop", "Belt Stop");
                validValues.Add("Reversing", "Reversing");
                validValues.Add("Air Blast", "Air Blast");
                validValues.Add("Kicker", "Kicker");
                validValues.Add("Diverter", "Diverter");
                validValues.Add("Retract", "Retract");
                validValues.Add("Manual", "Manual");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RejTyp1", "Reject Type", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("4", "4");
                validValues.Add("6", "6");
                validValues.Add("8", "8");
                validValues.Add("10", "10");
                validValues.Add("12", "12");
                validValues.Add("14", "14");
                validValues.Add("16", "16");
                validValues.Add("18", "18");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RejRet", "Reject Retract Type", 2, 2, BoFieldTypes.db_Alpha, null, string.Empty, validValues, "6");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_SecRej", "Second Reject Required", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                validValues = new Dictionary<string, string>();
                validValues.Add("Belt Stop", "Belt Stop");
                validValues.Add("Reversing", "Reversing");
                validValues.Add("Air Blast", "Air Blast");
                validValues.Add("Kicker", "Kicker");
                validValues.Add("Diverter", "Diverter");
                validValues.Add("Retract", "Retract");
                validValues.Add("Manual", "Manual");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RejTyp2", "Secondary Reject", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("Towards Controls", "Towards Controls");
                validValues.Add("Away From Controls", "Away From Controls");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RejDir", "Reject Direction", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, "Away From Controls");
                validValues = new Dictionary<string, string>();
                validValues.Add("None", "None");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RejAlrm", "Reject Alarm", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("None", "None");
                validValues.Add("Stainless Steel", "Stainless Steel");
                validValues.Add("Plastic", "Plastic");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RejBinTp", "Reject Bin Type", 20, 20, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                validValues = new Dictionary<string, string>();
                validValues.Add("None", "None");
                validValues.Add("Regal", "Regal");
                validValues.Add("Tray", "Tray");
                validValues.Add("Slide", "Slide");
                validValues.Add("Bakery", "Bakery");
                validValues.Add("Mac", "Mac");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_RejBinSt", "Reject Bin Style", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                // Belt Stop Alarm
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BSAPot", "BSA Potentiometer", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BSAKeyR", "BSA Key Reset", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BSABltR", "BSA Belt Reverse", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BSABtSw", "BSA Booted Switches", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_BSAEmSt", "BSA Emergency Stop", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                validValues = new Dictionary<string, string>();
                validValues.Add("Soft", "Soft");
                validValues.Add("Relay", "Relay");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_LatcType2", "Secondary Latch Type", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                // Air Blast
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_AirBTR", "Air Blast Tank Required", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                validValues = new Dictionary<string, string>();
                validValues.Add("2L", "2L");
                validValues.Add("10L", "10L");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_CapTank", "Air Blast Capacity Tank", 10, 10, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_AirPrS", "Air Pressure Sensor", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");

                // Kicker
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_KickHD", "Heavy Duty Kicker", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_KickFM", "Frame Mount Kicker Bracket", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");

                // Flap Gate
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_FlGtLn", "Flap Gate Overall Length", 0, 0, BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_FlChute", "Flap Gate Chute Required", 1, 1, BoFieldTypes.db_Alpha, null, string.Empty, yesNoList, "N");

                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_UNotes", "User Notes", 0, 0, BoFieldTypes.db_Memo, null, string.Empty, null, null);
                table.Fields.Add(addOn.Company, "XX_CONVBLT", "XX_SNotes", "User Notes", 0, 0, BoFieldTypes.db_Memo, null, string.Empty, null, null);


                // Provide Registration Informations
                table.Registration.ObjectName = "Conveyor Belt Configurator";
                table.Registration.ObjectCode = "CBConfigurator";
                table.Registration.CanCancel = true;
                table.Registration.CanClose = true;
                table.Registration.CanDelete = true;
                table.Registration.CanFind = true;
                table.Registration.CanLog = true;
                table.Registration.FindColumns.Add("Name");
                table.Registration.FindColumns.Add("Code");

                table.StatusbarChanged += addOn.StatusbarChanged;
                table.Add();
                table.StatusbarChanged -= addOn.StatusbarChanged;

                return true;

            }
            catch (Exception ex)
            {
                addOn.HandleException(ex);
                return false;
            }
        }

        /// <summary>
        /// Updates the sales order.
        /// </summary>
        /// <param name="addOn">The add on.</param>
        /// <returns>true, if successful</returns>
        public static bool UpdateSalesOrder(AddOn addOn)
        {
            using (var table = new TableAdapter(addOn.Company, "RDR1"))
            {
                table.Load();

                if (!table.Exists)
                {
                    return false;
                }
                try
                {
                    var validValues = new Dictionary<string, string>
                                          {
                                              {"No", "No"},
                                              {"Yes", "Yes"}
                                          };

                    table.Fields.Add(addOn.Company, "RDR1", "XX_ApprovalReq", "Approval Required", 50, 50, BoFieldTypes.db_Alpha, null, string.Empty, validValues, null);

                    table.StatusbarChanged += addOn.StatusbarChanged;
                    table.Update();
                    table.StatusbarChanged -= addOn.StatusbarChanged;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

    }
}
