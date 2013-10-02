//-----------------------------------------------------------------------
// <copyright file="FortressConfigurationHelper.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
//-----------------------------------------------------------------------

namespace Fortress.DI.Helpers
{

    #region Using Directive(s)

    using System.IO;
    using B1C.SAP.DI.Helpers;

    using SAPbobsCOM;
    using System;
    using System.Collections.Generic;

    #endregion Using Directive(s)

    public class FortressConfigurationHelper
    {

        /// <summary>
        /// Gets the sales order directory pattern.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static string GetSalesOrderDirectoryPattern(Company company)
        {
            string directoryPattern;
            ConfigurationHelper.GlobalConfiguration.Load(company, "SODirPtn", out directoryPattern);  

            return directoryPattern;
        }

        /// <summary>
        /// Gets the line directory pattern.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static string GetLineDirectoryPattern(Company company)
        {
            string directoryPattern;
            ConfigurationHelper.GlobalConfiguration.Load(company, "LnDirPtn", out directoryPattern);

            return directoryPattern;
        }

        /// <summary>
        /// Gets the bom file pattern.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static string GetBomFilePattern(Company company)
        {
            string directoryPattern;
            ConfigurationHelper.GlobalConfiguration.Load(company, "SpBomPtn", out directoryPattern);

            return directoryPattern;
        }

        /// <summary>
        /// Gets the directory in which the Production Order XML (.bom) files are stored
        /// </summary>
        public static DirectoryInfo GetDrawingDirectory(Company company)
        {
            string directory;
            ConfigurationHelper.GlobalConfiguration.Load(company, "SpBomDir", out directory);

            // Only load the directory if it is not null
            var drawingDirectory = new DirectoryInfo(directory);
            if (!drawingDirectory.Exists)
            {
                drawingDirectory = null;
            }

            return drawingDirectory;
        }

        /// <summary>
        /// Checks if we should auto-select the first row if none are selected
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns>true if the first row should be auto selected, false otherwise</returns>
        public static bool DoAutoSelectFirstRow(Company company)
        {
            bool select;
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "AutoSel", out select);
            }
            catch (Exception)
            {
                select = false;
            }            

            return select;
        }


        /// <summary>
        /// Gets the configuration notification user.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static IList<string> GetConfigurationNotificationUser(Company company)
        {
            string user;
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "ConfNtf", out user);
            }
            catch (Exception)
            {
                user = string.Empty;
            }

            IList<string> users = new List<string>();
            string[] splitUsers = user.Split(',', ';');

            for (int i = 0; i < splitUsers.Length; i++)
            {
                users.Add(splitUsers[i]);
            }
            
            return users;
        }

        /// <summary>
        /// Gets the configuration notification subject.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns>the subject for the configuration notifications</returns>
        public static string GetConfigurationNotificationSubject(Company company)
        {
            string value;
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "ConfSubj", "A configuration has been revised", out value);
            }
            catch (Exception)
            {
                value = "A configuration has been revised";
            }

            return value;
        }

        /// <summary>
        /// Gets the configuration notification body.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns>the body for the configuration notifications</returns>
        public static string GetConfigurationNotificationBody(Company company)
        {
            string value;
            const string defaultVal = "The configuration for Sales Order {OrderId}, line {LineNum} has been revised from revision {oldRev} to {newRev}";
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "ConfBody", defaultVal, out value);
            }
            catch (Exception)
            {
                value = defaultVal;
            }

            return value;
        }

        /// <summary>
        /// Gets the sales order directory initial.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static int GetSalesOrderDirectoryInitial(Company company)
        {
            int value;
            const int defaultVal = 10000;
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "SoDirSt", defaultVal, out value);
            }
            catch (Exception)
            {
                value = defaultVal;
            }

            return value;
        }


        /// <summary>
        /// Gets the sales order directory factor.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static int GetSalesOrderDirectoryFactor(Company company)
        {
            int value;
            const int defaultVal = 10;
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "SoDirFa", defaultVal, out value);
            }
            catch (Exception)
            {
                value = defaultVal;
            }

            return value;
        }

        /// <summary>
        /// Gets the sales order directory numeric format.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static string GetSalesOrderDirectoryNumericFormat(Company company)
        {
            string value;
            const string defaultVal = "00000";
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "SoDirFrm", defaultVal, out value);
            }
            catch (Exception)
            {
                value = defaultVal;
            }

            return value;
        }


        /// <summary>
        /// Gets the metal detector commision rate.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static double GetMetalDetectorCommisionRate(Company company)
        {
            double value;
            const double defaultVal = 0.15d;
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "MdComRt", out value);
            }
            catch (Exception)
            {
                value = defaultVal;
            }

            return value;
        }

        /// <summary>
        /// Gets the conveyor commision rate.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static double GetConveyorCommisionRate(Company company)
        {
            double value;
            const double defaultVal = 0.10d;
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "CnComRt", out value);
            }
            catch (Exception)
            {
                value = defaultVal;
            }

            return value;
        }

        /// <summary>
        /// Gets the conveyor commision rate.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static double GetPipelineCommisionRate(Company company)
        {
            double value;
            const double defaultVal = 0.10d;
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "PpComRt", out value);
            }
            catch (Exception)
            {
                value = defaultVal;
            }

            return value;
        }

        /// <summary>
        /// Gets the conveyor commision rate.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static double GetTabletSystemCommisionRate(Company company)
        {
            double value;
            const double defaultVal = 0.10d;
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "TbtComRt", out value);
            }
            catch (Exception)
            {
                value = defaultVal;
            }

            return value;
        }

        /// <summary>
        /// Gets the conveyor commision rate.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns></returns>
        public static double GetGravityRejectSystemCommisionRate(Company company)
        {
            double value;
            const double defaultVal = 0.10d;
            try
            {
                ConfigurationHelper.GlobalConfiguration.Load(company, "GrvComRt", out value);
            }
            catch (Exception)
            {
                value = defaultVal;
            }

            return value;
        }
    }
}
