//-----------------------------------------------------------------------
// <copyright file="ConfigurationHelper.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
// <summary>
//    Product        : B1C SAP Core Libraries
//    Author Email   : frank.alunni@B1C.com
//    Created Date   : 03/17/2009
//    Prerequisites  : - SAP Business One v.2007a
//                     - SQL Server 2005
//                     - Micorsoft .NET Framework 3.5
// </summary>
//-----------------------------------------------------------------------
namespace B1C.SAP.DI.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Windows.Forms;
    using BusinessAdapters;
    using B1C.SAP.DI.Business;

    /// <summary>
    /// Instance of the ConfigurationHelper class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public static class ConfigurationHelper
    {
        #region Fields
        /// <summary>
        /// System Configuration class
        /// </summary>
        private static Configuration configuration;
        #endregion Fields

        #region Methods

        /// <summary>
        /// Initializes this instance.
        /// </summary>
        public static void Initialize()
        {
            if (configuration == null)
            {
                configuration = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
            }
        }

        /// <summary>
        /// Loads the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="retVal">The return value.</param>
        public static void Load(string key, out string retVal)
        {
            Initialize();
            try
            {
                retVal = configuration.AppSettings.Settings[key].Value;
            }
            catch
            {
                retVal = string.Empty;
            }
        }

        /// <summary>
        /// Loads the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <param name="retVal">The return value.</param>
        public static void Load(string key, string defaultValue, out string retVal)
        {
            Initialize();
            try
            {
                retVal = configuration.AppSettings.Settings[key].Value;
            }
            catch
            {
                Save(key, defaultValue);
                retVal = defaultValue;
            }
        }

        /// <summary>
        /// Loads the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="retVal">The return value.</param>
        public static void Load(string key, out bool retVal)
        {
            Initialize();
            try
            {
                retVal = Convert.ToBoolean(configuration.AppSettings.Settings[key].Value);
            }
            catch
            {
                Save(key, false);
                retVal = false;
            }
        }

        /// <summary>
        /// Loads the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="retVal">The return value.</param>
        public static void Load(string key, out int retVal)
        {
            string returnValue;
            Load(key, out returnValue);
            int.TryParse(returnValue, out retVal);
        }

        /// <summary>
        /// Loads the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <param name="retVal">The return value.</param>
        public static void Load(string key, int defaultValue, out int retVal)
        {
            string returnValue;
            Load(key, defaultValue.ToString(), out returnValue);
            int.TryParse(returnValue, out retVal);
        }

        /// <summary>
        /// Loads the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="retVal">The return value.</param>
        public static void Load(string key, out float retVal)
        {
            string returnValue;
            Load(key, out returnValue);
            float.TryParse(returnValue, out retVal);
        }

        /// <summary>
        /// Loads the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="retVal">The return value.</param>
        public static void Load(string key, out double retVal)
        {
            string returnValue;
            Load(key, out returnValue);
            double.TryParse(returnValue, out retVal);
        }

        /// <summary>
        /// Loads the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="retVal">The return value.</param>
        public static void Load(string key, out decimal retVal)
        {
            string returnValue;
            Load(key, out returnValue);
            decimal.TryParse(returnValue, out retVal);
        }

        /// <summary>
        /// Loads the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="retVal">The return value.</param>
        public static void Load(string key, out DateTime retVal)
        {
            string returnValue;
            Load(key, out returnValue);
            DateTime.TryParse(returnValue, out retVal);
        }

        /// <summary>
        /// Saves the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="value">The element value.</param>
        public static void Save(string key, string value)
        {
            Initialize();
            try
            {
                configuration.AppSettings.Settings[key].Value = value;
            }
            catch
            {
                configuration.AppSettings.Settings.Add(key, value);
            }

            SaveConfig();
        }

        /// <summary>
        /// Saves the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="value">The element value.</param>
        public static void Save(string key, bool value)
        {
            Initialize();
            try
            {
                configuration.AppSettings.Settings[key].Value = value.ToString();
            }
            catch
            {
                configuration.AppSettings.Settings.Add(key, value.ToString());
            }

            SaveConfig();
        }

        /// <summary>
        /// Saves the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="value">The element value.</param>
        public static void Save(string key, int value)
        {
            Save(key, value.ToString());
        }

        /// <summary>
        /// Saves the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="value">The element value.</param>
        public static void Save(string key, double value)
        {
            Save(key, value.ToString());
        }

        /// <summary>
        /// Saves the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="value">The element value.</param>
        public static void Save(string key, decimal value)
        {
            Save(key, value.ToString());
        }

        /// <summary>
        /// Saves the specified key.
        /// </summary>
        /// <param name="key">The element key.</param>
        /// <param name="value">The element value.</param>
        public static void Save(string key, DateTime value)
        {
            Save(key, value.ToString());
        }

        /// <summary>
        /// Saves the config.
        /// </summary>
        private static void SaveConfig()
        {
            // Save the configuration file.
            configuration.Save(ConfigurationSaveMode.Modified);

            // Force a reload of the changed section.
            ConfigurationManager.RefreshSection("appSettings");
        }
        #endregion Methods

        /// <summary>
        /// Instance of the GlobalConfiguration class.
        /// </summary>
        /// <author>Frank Alunni</author>
        public static class GlobalConfiguration
        {
            /// <summary>
            /// A setting list
            /// </summary>
            private static readonly Dictionary<string, string> Settings = new Dictionary<string, string>();

            private static string _configurationTable;

            private static string _configurationValueColumn;

            private static string _configurationKeyColumn;

            /// <summary>
            /// Gets the configuration table.
            /// </summary>
            /// <value>The configuration table.</value>
            private static string ConfigurationTable
            {
                get
                {
                    if (string.IsNullOrEmpty(_configurationTable))
                    {
                        _configurationTable = ConfigurationManager.AppSettings["ConfigTable"];
                    }

                    return _configurationTable;
                }
            }

            /// <summary>
            /// Gets the configuration value column.
            /// </summary>
            /// <value>The configuration value column.</value>
            private static string ConfigurationValueColumn
            {
                get
                {
                    if (string.IsNullOrEmpty(_configurationValueColumn))
                    {
                        _configurationValueColumn = ConfigurationManager.AppSettings["ConfigValueColumn"];
                    }

                    return _configurationValueColumn;
                }
            }

            /// <summary>
            /// Gets the configuration key column.
            /// </summary>
            /// <value>The configuration key column.</value>
            private static string ConfigurationKeyColumn
            {
                get
                {
                    if (string.IsNullOrEmpty(_configurationKeyColumn))
                    {
                        _configurationKeyColumn = ConfigurationManager.AppSettings["ConfigKeyColumn"];
                    }

                    return _configurationKeyColumn;
                }
            }


            /// <summary>
            /// Loads the specified key.
            /// </summary>
            /// <param name="company">The company.</param>
            /// <param name="key">The element key.</param>
            /// <param name="retVal">The return value.</param>
            public static void Load(SAPbobsCOM.Company company, string key, out string retVal)
            {
                try
                {
                    using (var settingRecordSet = new RecordsetAdapter(company, string.Format("SELECT {0} FROM [{1}] WHERE {2} = '{3}'", ConfigurationValueColumn, ConfigurationTable, ConfigurationKeyColumn, key)))                        
                    {
                        retVal = settingRecordSet.EoF ? string.Empty : settingRecordSet.FieldValue(0).ToString();
                    }
                }
                catch
                {
                    retVal = string.Empty;
                }
            }

            /// <summary>
            /// Loads the specified key.
            /// </summary>
            /// <param name="company">The company.</param>
            /// <param name="key">The element key.</param>
            /// <param name="defaultValue">The default value.</param>
            /// <param name="retVal">The return value.</param>
            public static void Load(SAPbobsCOM.Company company, string key, string defaultValue, out string retVal)
            {
                try
                {
                    Load(company, key, out retVal);
                    if (string.IsNullOrEmpty(retVal))
                    {
                        retVal = defaultValue;
                    }
                }
                catch
                {
                    retVal = defaultValue;
                }
            }

            /// <summary>
            /// Loads the specified key.
            /// </summary>
            /// <param name="company">The company.</param>
            /// <param name="key">The element key.</param>
            /// <param name="retVal">The return value.</param>
            public static void Load(SAPbobsCOM.Company company, string key, out bool retVal)
            {
                try
                {
                    string value;
                    Load(company, key, out value);
                    retVal = Convert.ToBoolean(value);
                }
                catch
                {
                    retVal = false;
                }
            }

            /// <summary>
            /// Loads the specified key.
            /// </summary>
            /// <param name="company">The company.</param>
            /// <param name="key">The element key.</param>
            /// <param name="retVal">The return value.</param>
            public static void Load(SAPbobsCOM.Company company, string key, out int retVal)
            {
                try
                {
                    string value;
                    Load(company, key, out value);
                    retVal = Convert.ToInt32(value);
                }
                catch 
                {
                    retVal = 0;
                }
            }

            /// <summary>
            /// Loads the specified key.
            /// </summary>
            /// <param name="company">The company.</param>
            /// <param name="key">The element key.</param>
            /// <param name="defaultValue">The default value.</param>
            /// <param name="retVal">The return value.</param>
            public static void Load(SAPbobsCOM.Company company, string key, int defaultValue, out int retVal)
            {
                try
                {
                    string value;
                    Load(company, key, out value);
                    retVal = Convert.ToInt32(value);
                    if (retVal == 0)
                    {
                        retVal = defaultValue;
                    }
                }
                catch
                {
                    retVal = defaultValue;
                }
            }

            /// <summary>
            /// Loads the specified key.
            /// </summary>
            /// <param name="company">The company.</param>
            /// <param name="key">The element key.</param>
            /// <param name="retVal">The return value.</param>
            public static void Load(SAPbobsCOM.Company company, string key, out float retVal)
            {
                try
                {
                    string value;
                    Load(company, key, out value);
                    float.TryParse(value, out retVal);
                }
                catch
                {
                    retVal = 0;
                }
            }

            /// <summary>
            /// Loads the specified key.
            /// </summary>
            /// <param name="company">The company.</param>
            /// <param name="key">The element key.</param>
            /// <param name="retVal">The return value.</param>
            public static void Load(SAPbobsCOM.Company company, string key, out double retVal)
            {
                try
                {
                    string value;
                    Load(company, key, out value);
                    double.TryParse(value, out retVal);
                }
                catch
                {
                    retVal = 0;
                }
            }

            /// <summary>
            /// Loads the specified key.
            /// </summary>
            /// <param name="company">The company.</param>
            /// <param name="key">The element key.</param>
            /// <param name="retVal">The return value.</param>
            public static void Load(SAPbobsCOM.Company company, string key, out decimal retVal)
            {
                try
                {
                    string value;
                    Load(company, key, out value);
                    decimal.TryParse(value, out retVal);
                }
                catch
                {
                    retVal = 0;
                }
            }

            /// <summary>
            /// Loads the specified key.
            /// </summary>
            /// <param name="company">The company.</param>
            /// <param name="key">The element key.</param>
            /// <param name="retVal">The return value.</param>
            public static void Load(SAPbobsCOM.Company company, string key, out DateTime retVal)
            {
                try
                {
                    string value;
                    Load(company, key, out value);
                    DateTime.TryParse(value, out retVal);
                }
                catch
                {
                    retVal = DateTime.MinValue;
                }
            }

            /// <summary>
            /// Saveto the specified key.
            /// </summary>
            /// <param name="company">The company.</param>
            /// <param name="key">The element key.</param>
            /// <param name="newVal">The return value.</param>
            public static void Save(SAPbobsCOM.Company company, string key, bool newVal)
            {
                try
                {
                    Save(company, key, newVal.ToString());
                }
                catch
                {
                }
            }

            /// <summary>
            /// Saves to the specified key.
            /// </summary>
            /// <param name="company">The company.</param>
            /// <param name="key">The element key.</param>
            /// <param name="newVal">The new value.</param>
            public static void Save(SAPbobsCOM.Company company, string key, string newVal)
            {
                try
                {
                    using (var setting = new RecordsetAdapter(company, string.Format("SELECT 1 FROM [{0}] WHERE {1} = '{2}'", ConfigurationTable, ConfigurationKeyColumn, key)))
                    {
                        if (setting.EoF)
                        {
                            // Create Setting
                            var loadData = new System.Text.StringBuilder();
                            loadData.AppendLine("<DataLoader UDOName='GeneralSettings' UDOType='MasterData'>");
                            loadData.AppendLine("<Record>");
                            loadData.AppendFormat("<Field Name='Code' Type='System.String' Value='{0}' />", key);
                            loadData.AppendFormat("<Field Name='Name' Type='System.String' Value='{0}' />", key +  " description");
                            loadData.AppendFormat("<Field Name='U_XX_Value' Type='System.String' Value='{0}' />", newVal);
                            loadData.AppendLine("</Record>");
                            loadData.AppendLine("</DataLoader>");

                            var loader = DataLoader.InstanceFromXMLString(loadData.ToString());
                            if (loader != null)
                            {
                                loader.Process(company);

                                if (!string.IsNullOrEmpty(loader.ErrorMessage))
                                {
                                    throw new Exception("Data loader exception: " + loader.ErrorMessage);
                                }
                            }

                        }
                        else
                        {
                            // Update setting
                            using (var settingRecordSet = new RecordsetAdapter(company))
                            {
                                settingRecordSet.Execute(string.Format("UPDATE [{0}] SET {1} = '{2}' WHERE {3} = '{4}'", ConfigurationTable, ConfigurationValueColumn, newVal, ConfigurationKeyColumn, key));
                            }
                        }
                    }

                    if (Settings.ContainsKey(key))
                    {
                        Settings.Remove(key);
                    }
                }
                catch
                {
                }
            }
        }
    }
}
