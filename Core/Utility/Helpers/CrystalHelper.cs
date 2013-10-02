// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CrystalHelper.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   <p>The CrystalHelper class handles a number of actions that are usually
//   required to use Crystal Reports for Visual Studio .Net.
//   </p>
//   <p>This class can be used for all reports based on typed datasets and reports that
//   have embedded queries.</p>
// </summary>
// --------------------------------------------------------------------------------------------------------------------
namespace B1C.Utility.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Windows.Forms;

    using Components;

    // Crystalreports required assemblies
    using CrystalDecisions.CrystalReports.Engine;
    using CrystalDecisions.Shared;
    using CrystalDecisions.Windows.Forms;

    /// <summary>
    /// <p>The CrystalHelper class handles a number of actions that are usually
    /// required to use Crystal Reports for Visual Studio .Net. 
    /// </p>
    /// <p>This class can be used for all reports based on typed datasets and reports that
    /// have embedded queries.</p>
    /// </summary>
    /// <example>
    /// The following example shows how CrystalHelper can be used to print a report
    /// where the information is supplied by a DataSet:
    /// <code>
    /// void PrintReport(SqlConnection conn)
    /// {
    ///     using (DataSet ds = new TestData())
    ///     {
    ///         SqlHelper.FillDataset(conn, 
    ///         CommandType.Text, "SELECT * FROM Customers", 
    ///         ds, new string [] {"Customers"});
    ///         using (CrystalHelper helper = new CrystalHelper(new CrystalReport1()))
    ///         {
    ///             helper.DataSource = ds;
    ///             helper.Open();
    ///             helper.Print();
    ///             helper.Close();
    ///         }
    ///     }
    /// }
    /// </code>
    /// </example>
    public class CrystalHelper : IDisposable
    {
        #region Local Variables

        /// <summary>
        /// The parameter list
        /// </summary>
        private readonly SortedList<string, object> parameters;

        /// <summary>
        /// The parameter list
        /// </summary>
        private readonly SortedList<string, object> parametersList;

        /// <summary>
        /// The report document
        /// </summary>
        private ReportDocument reportDocument;

        /// <summary>
        /// The report data
        /// </summary>
        private DataSet reportData;

        /// <summary>
        /// The report file
        /// </summary>
        private string reportFile;

        /// <summary>
        /// Flag indicating wether the file is open or not
        /// </summary>
        private bool reportIsOpen;
        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CrystalHelper"/> class.
        /// </summary>
        public CrystalHelper()
        {
            this.parameters = new SortedList<string, object>();
            this.parametersList = new SortedList<string, object>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CrystalHelper"/> class.
        /// </summary>
        /// <param name="report">The <see cref="ReportDocument"/> object for an embedded report.</param>
        public CrystalHelper(ReportDocument report)
            : this()
        {
            this.reportDocument = report;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CrystalHelper"/> class.
        /// </summary>
        /// <param name="reportFile">Name and path for a CrystalReports (*.rpt) file.</param>
        public CrystalHelper(string reportFile)
            : this()
        {
            this.reportDocument = this.CreateReportDocument(reportFile);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CrystalHelper"/> class.
        /// </summary>
        /// <param name="report">The <see cref="ReportDocument"/> object for an embedded report.</param>
        /// <param name="serverName">Name of the database server.</param>
        /// <param name="databaseName">Name of the database.</param>
        /// <param name="userId">The user id required for logon.</param>
        /// <param name="userPassword">The password for the specified user.</param>
        public CrystalHelper(ReportDocument report, string serverName, string databaseName, string userId, string userPassword)
            : this()
        {
            this.reportDocument = report;

            // Setup the connection information 
            this.ServerName = serverName;
            this.DatabaseName = databaseName;
            this.UserId = userId;
            this.Password = userPassword;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CrystalHelper"/> class.
        /// </summary>
        /// <param name="report">The <see cref="ReportDocument"/> object for an embedded report.</param>
        /// <param name="serverName">Name of the database server.</param>
        /// <param name="databaseName">Name of the database.</param>
        /// <param name="integratedSecurity">if set to <c>true</c> integrated security is used to connect to the database; false if otherwise.</param>
        public CrystalHelper(ReportDocument report, string serverName, string databaseName, bool integratedSecurity)
            : this()
        {
            this.reportDocument = report;

            // Setup the connection information 
            this.ServerName = serverName;
            this.DatabaseName = databaseName;
            this.IntegratedSecurity = integratedSecurity;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CrystalHelper"/> class.
        /// </summary>
        /// <param name="reportFile">Name and path for a CrystalReports (*.rpt) file.</param>
        /// <param name="serverName">Name of the database server.</param>
        /// <param name="databaseName">Name of the database.</param>
        /// <param name="userId">The user id required for logon.</param>
        /// <param name="userPassword">The password for the specified user.</param>
        public CrystalHelper(string reportFile, string serverName, string databaseName, string userId, string userPassword)
            : this()
        {
            this.reportDocument = this.CreateReportDocument(reportFile);

            // Setup the connection information 
            this.ServerName = serverName;
            this.DatabaseName = databaseName;
            this.UserId = userId;
            this.Password = userPassword;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CrystalHelper"/> class.
        /// </summary>
        /// <param name="reportFile">Name and path for a CrystalReports (*.rpt) file.</param>
        /// <param name="serverName">Name of the database server.</param>
        /// <param name="databaseName">Name of the database.</param>
        /// <param name="integratedSecurity">if set to <c>true</c> integrated security is used to connect to the database; false if otherwise.</param>
        public CrystalHelper(string reportFile, string serverName, string databaseName, bool integratedSecurity)
            : this()
        {
            this.reportDocument = this.CreateReportDocument(reportFile);

            // Setup the connection information 
            this.ServerName = serverName;
            this.DatabaseName = databaseName;
            this.IntegratedSecurity = integratedSecurity;
        }

        #endregion

        #region Enums
        /// <summary>
        /// Supported export formats for Crystal reports.
        /// </summary>
        public enum CrystalExportFormat
        {
            /// <summary>
            /// Export to Microsoft Word (*.doc)
            /// </summary>
            Word,

            /// <summary>
            /// Export to Microsoft Excel (*.xls)
            /// </summary>
            Excel,

            /// <summary>
            /// Export to Rich Text format (*.rtf)
            /// </summary>
            RichText,

            /// <summary>
            /// Export to Adobe Portable Doc format (*.pdf)
            /// </summary>
            PortableDocFormat,

            /// <summary>
            /// Export to excel (*.xls)
            /// </summary>
            ExcelRecord,

            /// <summary>
            /// Export to html 32 (*.html)
            /// </summary>
            HTML32,

            /// <summary>
            /// Export to html 40 (*.html)
            /// </summary>
            HTML40,

            /// <summary>
            /// Export to tab separated
            /// </summary>
            TabSeperatedText,

            /// <summary>
            /// Export to text file
            /// </summary>
            Text,

            /// <summary>
            /// Export to xml
            /// </summary>
            Xml,

            /// <summary>
            /// Export not formatted
            /// </summary>
            NoFormat
        }
        #endregion Enums

        #region Public static properties

        /// <summary>
        /// Gets the export filter for a <see cref="SaveFileDialog"/>
        /// </summary>
        /// <value>The export filter.</value>
        public static string ExportFilter
        {
            get
            {
                var filter = new System.Text.StringBuilder();

                filter.Append("Rich text (*.rtf)|*.rtf|");
                filter.Append("Excel (*.xls)|*.xls|");
                filter.Append("Word (*.doc)|*.doc|");
                filter.Append("Portable Document Format (*.pdf)|*.pdf");

                return filter.ToString();
            }
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Gets or sets the report source.
        /// </summary>
        /// <value>The report source.</value>
        /// <remarks>
        /// Use this property when you want to show the report in the 
        /// CrystalReportsViewer control.</remarks>
        public ReportDocument ReportSource
        {
            get
            {
                return this.reportDocument;
            }

            set
            {
                if (this.reportDocument != null)
                {
                    this.reportDocument.Dispose();
                }

                this.reportDocument = value;
                if (value == null)
                {
                    this.reportFile = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the report file.
        /// </summary>
        /// <value>The report file.</value>
        public string ReportFile
        {
            get
            {
                return this.reportFile;
            }

            set
            {
                if (this.reportDocument != null)
                {
                    this.reportDocument.Dispose();
                }

                if (value != null && value.Trim().Length > 0)
                {
                    this.reportDocument = this.CreateReportDocument(value);
                }
                else
                {
                    this.reportFile = value;
                    this.reportDocument = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the data source.
        /// </summary>
        /// <value>The data source.</value>
        public DataSet DataSource
        {
            get
            {
                return this.reportData;
            }

            set
            {
                if (this.reportData != null)
                {
                    this.reportData.Dispose();
                }

                this.reportData = value;
            }
        }

        /// <summary>
        /// Gets or sets the name of the server.
        /// </summary>
        /// <value>The name of the server.</value>
        public string ServerName { get; set; }

        /// <summary>
        /// Gets or sets the name of the database.
        /// </summary>
        /// <value>The name of the database.</value>
        public string DatabaseName { get; set; }

        /// <summary>
        /// Gets or sets the user id.
        /// </summary>
        /// <value>The user id.</value>
        public string UserId { get; set; }

        /// <summary>
        /// Gets or sets the password.
        /// </summary>
        /// <value>The password.</value>
        public string Password { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether [integrated security].
        /// </summary>
        /// <value><c>true</c> if integrated security is used to connect to the database; false if otherwise.</value>
        public bool IntegratedSecurity { get; set; }
        #endregion

        #region Public static methods

        /// <summary>
        /// Show or hide the status bar on the crystalreports viewer.
        /// </summary>
        /// <param name="viewer">Current viewer.</param>
        /// <param name="visible">Set if the status bar is visible or not.</param>
        /// <remarks>
        /// This method will throw an exception when the viewer object is not specified.
        /// </remarks>
        /// <example>
        /// This example shows how the status bar can be removed:
        /// <code>
        /// CrystalHelper.ViewerStatusBar(myViewer, false);
        /// </code>
        /// </example>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters", Justification = "Deliberately targeting viewer.")]
        public static void ViewerStatusBar(CrystalReportViewer viewer, bool visible)
        {
            if (viewer == null)
            {
                throw new ArgumentNullException("viewer");
            }

            foreach (Control control in viewer.Controls)
            {
                if (string.Compare(control.GetType().Name, "StatusBar", true, System.Globalization.CultureInfo.InvariantCulture) == 0)
                {
                    control.Visible = visible;
                }
            }
        }

        /// <summary>
        /// Show or hide the tabs on the crystalreports viewer.
        /// </summary>
        /// <param name="viewer">Current viewer</param>
        /// <param name="visible">Set if the status bar is visible or not.</param>
        /// <remarks>
        /// This method will throw an exception when the viewer object is not specified.
        /// </remarks>
        /// <example>
        /// This example shows how the tabs can be removed:
        /// <code>
        /// CrystalHelper.ViewerTabs(myViewer, false);
        /// </code>
        /// </example>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters", Justification = "Deliberately targeting viewer.")]
        public static void ViewerTabs(CrystalReportViewer viewer, bool visible)
        {
            if (viewer == null)
            {
                throw new ArgumentNullException("viewer");
            }

            foreach (Control control in viewer.Controls)
            {
                if (string.Compare(control.GetType().Name, "PageView", true, System.Globalization.CultureInfo.InvariantCulture) == 0)
                {
                    var tab = (TabControl)control.Controls[0];

                    if (!visible)
                    {
                        tab.ItemSize = new System.Drawing.Size(0, 1);
                        tab.SizeMode = TabSizeMode.Fixed;
                        tab.Appearance = TabAppearance.Buttons;
                    }
                    else
                    {
                        tab.ItemSize = new System.Drawing.Size(67, 18);
                        tab.SizeMode = TabSizeMode.Normal;
                        tab.Appearance = TabAppearance.Normal;
                    }
                }
            }
        }

        /// <summary>
        /// Replace the current name of a tab with a new name.
        /// </summary>
        /// <param name="viewer">Current viewer.</param>
        /// <param name="oldName">The name to be replaced.</param>
        /// <param name="newName">The new name.</param>
        /// <remarks>
        /// This method will throw an exception when:
        /// <list type="bullet">
        /// <item><description>The viewer object is not specified.</description></item>
        /// <item><description>The Open() method is not yet called on this object.</description></item>
        /// </list>
        /// </remarks>
        /// <example>
        /// This example shows how the default tab description MainReport can be replaced.
        /// <code>
        /// CrystalHelper.ViewerStatusBar(myViewer, "MainReport", "And now for something completely different");
        /// </code>
        /// </example>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters", Justification = "Deliberately targeting viewer.")]
        public static void ReplaceReportName(CrystalReportViewer viewer, string oldName, string newName)
        {
            if (viewer == null)
            {
                throw new ArgumentNullException("viewer");
            }

            if (string.IsNullOrEmpty(oldName))
            {
                throw new ArgumentException("May not be empty.", "oldName");
            }

            foreach (Control control in viewer.Controls)
            {
                if (string.Compare(control.GetType().Name, "PageView", true, System.Globalization.CultureInfo.InvariantCulture) == 0)
                {
                    foreach (Control controlInPage in control.Controls)
                    {
                        if (string.Compare(controlInPage.GetType().Name, "TabControl", true, System.Globalization.CultureInfo.InvariantCulture) == 0)
                        {
                            var tabs = (TabControl)controlInPage;

                            foreach (TabPage tabPage in tabs.TabPages)
                            {
                                if (string.Compare(tabPage.Text, oldName, false, System.Globalization.CultureInfo.InvariantCulture) == 0)
                                {
                                    tabPage.Text = newName;
                                    return;
                                }
                            }
                        }
                    }
                }
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Set value for report parameters
        /// </summary>
        /// <param name="name">Parameter name.</param>
        /// <param name="value">Value to be set for the specified parameter.</param>
        public void SetParameterList(string name, object value)
        {
            if (this.parametersList.ContainsKey(name))
            {
                this.parametersList[name] = value;
            }
            else
            {
                this.parametersList.Add(name, value);
            }
        }

        /// <summary>
        /// Set value for report parameters
        /// </summary>
        /// <param name="name">Parameter name.</param>
        /// <param name="value">Value to be set for the specified parameter.</param>
        public void SetParameter(string name, object value)
        {
            if (this.parameters.ContainsKey(name))
            {
                this.parameters[name] = value;
            }
            else
            {
                this.parameters.Add(name, value);
            }
        }

        /// <summary>
        /// Close the report.
        /// </summary>
        /// <remarks>
        /// This method will throw an exception when the Open() method is not yet called on this object.
        /// </remarks>
        public void Close()
        {
            if (!this.reportIsOpen)
            {
                throw new InvalidOperationException("The report is already closed.");
            }

            this.reportDocument.Close();
            this.reportIsOpen = false;
        }

        /// <summary>
        /// Open the report.
        /// </summary>
        /// <remarks>
        /// <p>This method will first attempt to assign the DataSource if that was specified. When no DataSource has been
        /// assigned, a check will be made to see if database connection information has been specified. When this is the case
        /// this information will be assigned to the report.
        /// </p>
        /// <P>This method will throw an exception when:
        /// <list type="bullet">
        /// <item><description>The report (rpt) has not been assigned yet.</description></item>
        /// <item><description>The Open() method has already been called on this object.</description></item>
        /// <item><description>A table being used in the report does not exist in the dataset which has been assigned to this report. (only when a dataset has been assigned to the DataSource property)</description></item>
        /// <item><description>The ServerName, DatabaseName or UserId property has not been set. (only when the DataSource property has not been set)</description></item>
        /// <item><description>When no database connection or datset could be assignd.</description></item>
        /// </list>
        /// </P></remarks>
        public void Open()
        {
            if (this.reportDocument == null)
            {
                throw new CrystalHelperException("First assign a report document.");
            }

            if (this.reportIsOpen)
            {
                throw new InvalidOperationException("The report is already open.");
            }

            this.SetParameters();

            // Check if the connection object exists. If so assign that
            // to the report.
            if (this.reportData != null)
            {
                // Assign the dataset to the report.
                this.AssignDataSet();
            }
            else
            {
                if (this.ServerName.Length == 0 || (!this.IntegratedSecurity && this.UserId.Length == 0) || this.DatabaseName.Length == 0)
                {
                    throw new CrystalHelperException("Connection information is incomplete. Report could not be opened.");
                }

                this.AssignConnection();
            }

            this.reportIsOpen = true;
        }

        /// <summary>
        /// Force a refresh of the data in the report. 
        /// </summary>
        /// <remarks>
        /// When the report is based on a DataSource, the report will be refreshed using 
        /// data in that DataSource. In case the report has it's own database connection 
        /// and uses SQL queries, the report will be refreshed using that information. 
        /// <br/>
        /// <p>
        /// This method will throw an exception when the Open() method is not yet called on this object.
        /// </p>
        /// </remarks>
        public void Refresh()
        {
            if (!this.reportIsOpen)
            {
                throw new InvalidOperationException("The report is not open.");
            }

            this.reportDocument.Refresh();
        }

        /// <summary>
        /// Print the report to the specified printer.
        /// </summary>
        /// <param name="printerName">Name of the printer to print the information to.</param>
        /// <param name="numberOfCopies">The number of copies.</param>
        /// <param name="collatePages">Indicates whether to collate the pages.</param>
        /// <param name="firstPage">First page to be printed. When less than 1, printing starts at the first page.</param>
        /// <param name="lastPage">Last page to be printed. When less than 1, or more than 9999, the last page will be printed as last page.</param>
        /// <remarks>
        /// This method will throw an exception when the Open() method is not yet called on this object.
        /// </remarks>
        public void Print(string printerName, int numberOfCopies, bool collatePages, int firstPage, int lastPage)
        {
            bool openedHere = false;

            if (!this.reportIsOpen)
            {
                this.Open();
                openedHere = true;
            }

            this.reportDocument.PrintOptions.PrinterName = printerName;

            if (firstPage < 1)
            {
                firstPage = 0;
            }

            if ((lastPage < 1) || (lastPage > 9999))
            {
                lastPage = 0;
            }

            this.reportDocument.PrintToPrinter(numberOfCopies, collatePages, firstPage, lastPage);

            if (openedHere)
            {
                this.Close();
            }
        }

        /// <summary>
        /// Prints one copy of the entire report to the default printer.
        /// </summary>
        /// <remarks>
        /// This method will throw an exception when the Open() method is not yet called on this object.
        /// </remarks>
        public void Print()
        {
            this.Print(string.Empty, 1, false, 0, 0);
        }

        /// <summary>
        /// Prints one copy of the entire report to the default printer.
        /// </summary>
        /// <param name="numberOfCopies">The number of copies.</param>
        /// <remarks>
        /// This method will throw an exception when the Open() method is not yet called on this object.
        /// </remarks>
        public void Print(int numberOfCopies)
        {
            this.Print(string.Empty, numberOfCopies, false, 0, 0);
        }

        /// <summary>
        /// Prints the specified number of copies.
        /// </summary>
        /// <param name="numberOfCopies">The number of copies.</param>
        /// <param name="firstPage">The first page.</param>
        /// <param name="lastPage">The last page.</param>
        /// <remarks>
        /// This method will throw an exception when the Open() method is not yet called on this object.
        /// </remarks>
        public void Print(int numberOfCopies, int firstPage, int lastPage)
        {
            this.Print(string.Empty, numberOfCopies, false, firstPage, lastPage);
        }


        /// <summary>
        /// Export the report to a file. 
        /// </summary>
        /// <param name="fileName">Name of the export file.</param>
        /// <remarks>
        /// The type of document is specified by the extension for the file 
        /// name When no extension is specified, the export format will default to Rich text.
        /// The following document types are supported:
        /// <list type="bullet">
        /// <item><description>Word (*.doc).</description></item>
        /// <item><description>Excel (*.xls).</description></item>
        /// <item><description>Rich text (*.rtf).</description></item>
        /// <item><description>Portable Doc Format (*.pdf)</description></item>
        /// </list><br/>
        /// <p>
        /// This method will throw an exception when the Open() method is not yet called on this object.
        /// </p>
        /// </remarks>
        public void Export(string fileName)
        {
            this.Export(fileName, GetExportFormat(fileName));
        }

        /// <summary>
        /// Export the report to a file using the specified format.
        /// </summary>
        /// <param name="fileName">Name of the export file.</param>
        /// <param name="exportFormat">Crystal Export Format.</param>
        /// <remarks>
        /// The following document types are supported:
        /// <list type="bullet">
        ///        <item><description>Word (*.doc).</description></item>
        ///        <item><description>Excel (*.xls).</description></item>
        ///        <item><description>Rich text (*.rtf).</description></item>
        ///        <item><description>Portable Doc Format (*.pdf)</description></item>
        /// </list><br/>
        /// <p>
        /// This method will throw an exception when the Open() method is not yet called on this object.
        /// </p>
        /// </remarks>
        public void Export(string fileName, CrystalExportFormat exportFormat)
        {
            bool openedHere = false;

            if (!this.reportIsOpen)
            {
                this.Open();
                openedHere = true;
            }
            
            this.reportDocument.ExportToDisk(GetExportType(exportFormat), fileName);

            if (openedHere)
            {
                this.Close();
            }
        }
        #endregion

        #region IDisposable Members

        /// <summary>
        /// Dispose of this object's resources.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(true); // as a service to those who might inherit from us
        }

        /// <summary>
        /// Free the instance variables of this object.
        /// </summary>
        /// <param name="disposing">if set to <c>true</c> [disposing].</param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (this.reportIsOpen)
                {
                    this.reportDocument.Close();
                }

                if (this.reportDocument != null)
                {
                    this.reportDocument.Dispose();
                    this.reportDocument = null;
                }

                if (this.reportData != null)
                {
                    this.reportData.Dispose();
                    this.reportData = null;
                }
            }
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Get the Crystal Export format using the CrystalExportFormat export format definitions
        /// </summary>
        /// <param name="exportFormat"><see cref="CrystalExportFormat"/> export type</param>
        /// <returns>Crystal Reports <see cref="ExportFormatType"/> type</returns>
        private static ExportFormatType GetExportType(CrystalExportFormat exportFormat)
        {
            ExportFormatType result;

            switch (exportFormat)
            {
                case CrystalExportFormat.Word:
                    result = ExportFormatType.WordForWindows;
                    break;
                case CrystalExportFormat.Excel:
                    result = ExportFormatType.Excel;
                    break;
                case CrystalExportFormat.RichText:
                    result = ExportFormatType.RichText;
                    break;
                case CrystalExportFormat.PortableDocFormat:
                    result = ExportFormatType.PortableDocFormat;
                    break;
                case CrystalExportFormat.ExcelRecord:
                    result = ExportFormatType.ExcelRecord;
                    break;
                case CrystalExportFormat.HTML32:
                    result = ExportFormatType.HTML32;
                    break;
                case CrystalExportFormat.HTML40:
                    result = ExportFormatType.HTML40;
                    break;
                    /* These are full crystal version only
                case CrystalExportFormat.TabSeperatedText:
                    result = ExportFormatType.TabSeperatedText;
                    break;
                case CrystalExportFormat.Text:
                    result = ExportFormatType.Text;
                    break;
                case CrystalExportFormat.Xml:
                    result = ExportFormatType.Xml;
                    break;
                     */
                case CrystalExportFormat.NoFormat:
                    result = ExportFormatType.NoFormat;
                    break;
                default:
                    throw new CrystalHelperException(string.Format(System.Globalization.CultureInfo.CurrentUICulture, "Unsupported export format '{0}'.", exportFormat));
            }

            return result;
        }

        /// <summary>
        /// Get the CrystalExportFormat export format definitions using the file name extension
        /// </summary>
        /// <param name="fileName">Name of the export file</param>
        /// <returns><see cref="CrystalExportFormat"/> export type</returns>
        private static CrystalExportFormat GetExportFormat(string fileName)
        {
            CrystalExportFormat result;

            string extension = FileHelper.FileExtension(fileName).ToUpper(System.Globalization.CultureInfo.CurrentCulture);

            switch (extension)
            {
                case "DOC":
                    result = CrystalExportFormat.Word;
                    break;
                case "XLS":
                    result = CrystalExportFormat.Excel;
                    break;
                case "RTF":
                    result = CrystalExportFormat.RichText;
                    break;
                case "PDF":
                    result = CrystalExportFormat.PortableDocFormat;
                    break;
                default:
                    throw new CrystalHelperException(string.Format(System.Globalization.CultureInfo.CurrentUICulture, "Unsupported export format for file '{0}'.", fileName));
            }

            return result;
        }

        /// <summary>
        /// Assign a <see cref="ConnectionInfo"/> to a <see cref="CrystalDecisions.CrystalReports.Engine.Table"/> object
        /// </summary>
        /// <param name="table">Table to which the connection is to be assigned</param>
        /// <param name="connection">Connection to the database.</param>
        private static void AssignTableConnection(Table table, ConnectionInfo connection)
        {
            // Cache the logon info block
            TableLogOnInfo logOnInfo = table.LogOnInfo;

            // Set the connection
            logOnInfo.ConnectionInfo = connection;

            // Apply the connection to the table!
            table.ApplyLogOnInfo(logOnInfo);
        }

        /// <summary>
        /// Assign the database connection to the table in all the report sections.
        /// </summary>
        private void AssignConnection()
        {
            var connection = new ConnectionInfo
            {
                DatabaseName = this.DatabaseName,
                ServerName = this.ServerName
            };

            if (this.IntegratedSecurity)
            {
                connection.IntegratedSecurity = this.IntegratedSecurity;
            }
            else
            {
                connection.UserID = this.UserId;
                connection.Password = this.Password;
            }

            // First we assign the connection to all tables in the main report
            foreach (Table table in this.reportDocument.Database.Tables)
            {
                AssignTableConnection(table, connection);
            }

            // Now loop through all the sections and its objects to do the same for the subreports
            foreach (Section section in this.reportDocument.ReportDefinition.Sections)
            {
                // In each section we need to loop through all the reporting objects
                foreach (ReportObject reportObject in section.ReportObjects)
                {
                    if (reportObject.Kind == ReportObjectKind.SubreportObject)
                    {
                        var subReport = (SubreportObject)reportObject;
                        var subDocument = subReport.OpenSubreport(subReport.SubreportName);

                        foreach (Table table in subDocument.Database.Tables)
                        {
                            AssignTableConnection(table, connection);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Assign the DataSet to the report.
        /// </summary>
        private void AssignDataSet()
        {
            DataSet reportDataSet = this.reportData.Copy();

            // Remove primary key info. CR9 does not appreciate this information!!!
            foreach (DataTable dataTable in reportDataSet.Tables)
            {
                if ((dataTable != null) && (dataTable.PrimaryKey != null))
                {
                    foreach (var dataColumn in dataTable.PrimaryKey)
                    {
                        dataColumn.AutoIncrement = false;
                    }

                    dataTable.PrimaryKey = null;
                }
            }

            // Now assign the dataset to all tables in the main report
            this.reportDocument.SetDataSource(this.reportData);

            // Now loop through all the sections and its objects to do the same for the subreports
            foreach (Section section in this.reportDocument.ReportDefinition.Sections)
            {
                // In each section we need to loop through all the reporting objects
                foreach (ReportObject reportObject in section.ReportObjects)
                {
                    if (reportObject.Kind == ReportObjectKind.SubreportObject)
                    {
                        var subReport = (SubreportObject)reportObject;
                        var subDocument = subReport.OpenSubreport(subReport.SubreportName);

                        subDocument.SetDataSource(reportDataSet);
                    }
                }
            }
        }

        /// <summary>
        /// Create a ReportDocument object using a report file and store
        /// the name of the file.
        /// </summary>
        /// <param name="reportDocumentFile">Name of the CrystalReports file (*.rpt).</param>
        /// <returns>A valid ReportDocument object.</returns>
        private ReportDocument CreateReportDocument(string reportDocumentFile)
        {
            var newDocument = new ReportDocument();

            this.reportFile = reportDocumentFile;
            newDocument.Load(reportDocumentFile);

            return newDocument;
        }

        /// <summary>
        /// Sets the parameters that have been added using the SetParameter method
        /// </summary>
        private void SetParameters()
        {
            foreach (ParameterFieldDefinition parameter in this.reportDocument.DataDefinition.ParameterFields)
            {
                try
                {
                    // Now get the current value for the parameter
                    var currentValues = parameter.CurrentValues;
                    currentValues.Clear();

                    // Create a value object for Crystal reports and assign the specified value.
                    ParameterDiscreteValue newValue;

                    if (this.parameters.ContainsKey(parameter.Name))
                    {
                        newValue = new ParameterDiscreteValue
                                       {
                                           Value = this.parameters[parameter.Name]
                                       };

                        // Now add the new value to the values collection and apply the 
                        // collection to the report.
                        currentValues.Add(newValue);
                        parameter.ApplyCurrentValues(currentValues);
                    }
                    else
                    {
                        // Check if the parameter is a list of values
                        if (this.parametersList.ContainsKey(parameter.Name))
                        {
                            string[] values = this.parametersList[parameter.Name].ToString().Split(',');

                            foreach (string value in values)
                            {
                                newValue = new ParameterDiscreteValue
                                               {
                                                   Value = value
                                               };

                                // Now add the new value to the values collection and apply the 
                                // collection to the report.
                                currentValues.Add(newValue);
                            }

                            parameter.ApplyCurrentValues(currentValues);
                        }
                    }
                }
                catch
                {
                    // Ignore any errors
                }
            }
        }

        /// <summary>
        /// Sets the parameters that have been added using the SetParameter method
        /// </summary>
        public void SetViewerParameters(CrystalReportViewer viewer)
        {

            var source = (ReportDocument)viewer.ReportSource;

            foreach (ParameterFieldDefinition parameter in source.DataDefinition.ParameterFields)
            {
                try
                {
                    // Now get the current value for the parameter
                    var currentValues = parameter.CurrentValues;
                    currentValues.Clear();

                    // Create a value object for Crystal reports and assign the specified value.
                    ParameterDiscreteValue newValue;

                    if (this.parameters.ContainsKey(parameter.Name))
                    {
                        newValue = new ParameterDiscreteValue
                        {
                            Value = this.parameters[parameter.Name]
                        };

                        // Now add the new value to the values collection and apply the 
                        // collection to the report.
                        currentValues.Add(newValue);
                        parameter.ApplyCurrentValues(currentValues);
                    }
                    else
                    {
                        // Check if the parameter is a list of values
                        if (this.parametersList.ContainsKey(parameter.Name))
                        {
                            string[] values = this.parametersList[parameter.Name].ToString().Split(',');

                            foreach (string value in values)
                            {
                                newValue = new ParameterDiscreteValue
                                {
                                    Value = value
                                };

                                // Now add the new value to the values collection and apply the 
                                // collection to the report.
                                currentValues.Add(newValue);
                            }

                            parameter.ApplyCurrentValues(currentValues);
                        }
                    }
                }
                catch
                {
                    // Ignore any errors
                }
            }
        }

        #endregion
    }
}
