//-----------------------------------------------------------------------
// <copyright file="DocumentAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters
{
    #region Using Directive(s)

    using System;
    using System.Collections.Generic;
    using System.IO;
    using BusinessPartners;
    using Components;
    using Constants;
    using Helpers;
    using SAPbobsCOM;
    using Utility.Helpers;
    using Utility.Logging;
    using Model.Exceptions;
    using B1C.SAP.DI.BusinessAdapters.Inventory;

    #endregion Using Directive(s)

    /// <summary>
    /// Instance of the DocumentAdapter class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public abstract class DocumentAdapter : SapObjectAdapter
    {
        #region Fields
        /// <summary>
        /// A Collection of document lines
        /// </summary>
        private Lines _lines;

        /// <summary>
        /// The document business partner
        /// </summary>
        private BusinessPartnerAdapter _businessPartner;

        /// <summary>
        /// The originating draft 
        /// </summary>
        private int _draftId = -1;
        #endregion Fields

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentAdapter"/> class, with a blank (new, unsaved) document 
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        protected DocumentAdapter(Company company)
            : base(company)
        {
            this.Document = (Documents)this.Company.GetBusinessObject(this.DocumentType);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="document">The document</param>
        protected DocumentAdapter(Company company, Documents document)
            : base(company)
        {
            this.Document = document;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentAdapter"/> class.
        /// </summary>
        /// <param name="company">The SAP Company object</param>
        /// <param name="docEntry">The document id</param>
        protected DocumentAdapter(Company company, int docEntry)
            : base(company)
        {
            this.Document = this.LoadDocument(docEntry);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="xmlFile">The XML file.</param>
        /// <param name="index">The index of the object to load from the XML.</param>
        protected DocumentAdapter(Company company, string xmlFile, int index)
            : base(company)
        {
            try
            {
                this.Document = (Documents) company.GetBusinessObjectFromXML(xmlFile, index);
            }
            catch (Exception ex)
            {
                throw new SapImportException(ex);
            }
        }

        #endregion Constructor(s)

        #region Properties

        /// <summary>
        /// Gets or sets the base adapters.
        /// </summary>
        /// <value>The base adapters.</value>
        protected IDictionary<int, DocumentAdapter> BaseAdapters { get; set; }

        /// <summary>
        /// Gets or sets the document.
        /// </summary>
        /// <value>The document.</value>
        public Documents Document { get; protected set; }

        /// <summary>
        /// Gets the lines.
        /// </summary>
        /// <value>The lines.</value>
        public Lines Lines
        {
            get
            {
                if (this.Document == null)
                {
                    return null;
                }

                return this._lines ?? (this._lines = new Lines(this.Document.Lines));
            }
        }

        /// <summary>
        /// Gets or sets the base adapter.
        /// </summary>
        /// <value>The base adapter.</value>
        public DocumentAdapter BaseAdapter { get; set; }

        /// <summary>
        /// Gets the type of document this is an adapter for
        /// </summary>
        public abstract BoObjectTypes DocumentType { get; }

        /// <summary>
        /// Gets the target document type of this adapter's document type
        /// </summary>
        public abstract BoObjectTypes TargetType { get; }

        /// <summary>
        /// Gets the business partner.
        /// </summary>
        /// <value>The business partner.</value>
        public BusinessPartnerAdapter BusinessPartner
        {
            get 
            {
                return this._businessPartner ??
                       (this._businessPartner = new BusinessPartnerAdapter(this.Company, this.Document.CardCode));
            }
        }

        /// <summary>
        /// Gets the draft id.
        /// </summary>
        /// <value>The draft id.</value>
        public int DraftId
        {
            get
            {
                if (this._draftId == -1)
                {
                    try
                    {
                        using (var rs = new RecordsetAdapter(
                            this.Company, 
                            string.Format("SELECT IsNull(draftkey, 0) as Draftkey FROM [{0}] WHERE DocEntry = {1}", this.DocumentTable, this.Document.DocEntry)))
                        {
                            this._draftId = rs.EoF ? 0 : Convert.ToInt32(rs.FieldValue("DraftKey").ToString());
                        }
                    }
                    catch
                    {
                        this._draftId = 0;
                    }
                }

                return this._draftId;
            }
        }

        /// <summary>
        /// Gets the user field count.
        /// </summary>
        /// <value>The user field count.</value>
        public int UserFieldCount
        {
            get
            {
                int count;
                var ufs = this.Document.UserFields;
                try
                {
                    var fields = ufs.Fields;

                    try
                    {
                        count = fields.Count;
                    }
                    finally
                    {
                        COMHelper.Release(ref fields);
                    }
                }
                finally
                {
                    COMHelper.Release(ref ufs);
                }

                return count;
            }
        }

        /// <summary>
        /// Gets the document master table
        /// </summary>
        protected string DocumentTable
        {
            get
            {
                return SapDocumentTables.GetDocumentTable(this.DocumentType);
            }
        }

        /// <summary>
        /// Gets the document row table
        /// </summary>
        protected string DocumentRowTable
        {
            get
            {
                return SapDocumentTables.GetDocumentRowTable(this.DocumentType);
            }
        }

        /// <summary>
        /// Gets the document code
        /// </summary>
        protected int DocumentCode
        {
            get
            {
                return SapDocumentTables.GetDocumentCode(this.DocumentType);
            }
        }

        /// <summary>
        /// Gets the name of the report.
        /// </summary>
        /// <value>The name of the report.</value>
        protected string ReportName
        {
            get
            {
                return SapDocumentTables.GetReportName(this.DocumentType) + ".rpt";
            }
        }

        /// <summary>
        /// Gets the doc entry.
        /// </summary>
        public int DocEntry
        {
            get
            {
                return this.Document.DocEntry;
            }
        }

        /// <summary>
        /// Gets the doc number.
        /// </summary>
        public int DocNumber
        {
            get
            {
                return this.Document.DocNum;
            }
        }
        #endregion Properties

        #region Methods

        /// <summary>
        /// Adds this instance.
        /// </summary>
        /// <exception cref="SapException">Thrown if there is an SAP error adding the Document</exception>
        /// <returns>True of the method adds successfully</returns>
        public virtual bool Add()
        {
            this.Task = "Adding Document of type " + this.DocumentCode;
            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the DocumentAdapter.Add is not set.");
            }

            if (this.Document == null)
            {
                throw new NullReferenceException("The document has not been set. DocumentAdapter.Add");
            }

            int counter = 0;
            SapException ex = null;
            while (counter < 5)
            {
                if (this.Document.Add() == 0)
                {
                    ex = null;
                    break;
                }

                ex = new SapException(this.Company);

                if (!ex.DoRetry)
                {
                    break;
                }

                this.Task = string.Format("System busy, the document cannot be added at this time. Attempt {0}", counter + 1);
                System.Threading.Thread.Sleep(1000);
                counter++;
            }

            if (ex != null)
            {
                throw ex;
            }

            string newDocEntry = this.Company.GetNewObjectKey();

            ThreadedAppLog.WriteLine("Document successfully created with doc entry {0}", newDocEntry);

            int docEntry = Convert.ToInt32(newDocEntry);
            this.Document = this.LoadDocument(docEntry);

            return true;
        }

        /// <summary>
        /// Closes the document.
        /// </summary>
        /// <exception cref="SapException">Thrown if there is an SAP error closing the Document</exception>
        public virtual void Close()
        {
            this.Task = "Closing Document [" + this.DocumentCode + ": " + this.Document.DocEntry + "]";
            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the DocumentAdapter.Add is not set.");
            }

            if (this.Document == null)
            {
                throw new NullReferenceException("The document has not been set. DocumentAdapter.Add");
            }

            int counter = 0;
            SapException ex = null;
            while (counter < 5)
            {
                if (this.Document.Close() == 0)
                {
                    ex = null;
                    break;
                }

                ex = new SapException(this.Company);

                if (!ex.DoRetry)
                {
                    break;
                }

                this.Task = string.Format("System busy, the document cannot be closed at this time. Attempt {0}", counter + 1);
                System.Threading.Thread.Sleep(1000);
                counter++;
            }

            if (ex != null)
            {
                throw ex;
            }
       }

        /// <summary>
        /// Updates this instance.
        /// </summary>
        /// <exception cref="SapException">Thrown if there is an SAP error updating the Document</exception>
        /// <returns>True of the method updates successfully</returns>
        public bool Update()
        {
            this.Task = "Updating Document [" + this.DocumentCode + ": " + this.Document.DocEntry + "]";

            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the DocumentAdapter.Update is not set.");
            }

            if (this.Document == null)
            {
                throw new NullReferenceException("The document has not been set. DocumentAdapter.Update");
            }


            int counter = 0;
            SapException ex = null;
            while (counter < 5)
            {
                if (this.Document.Update() == 0)
                {
                    ex = null;
                    break;
                }

                ex = new SapException(this.Company);

                if (!ex.DoRetry)
                {
                    break;
                }

                this.Task = string.Format("System busy, the document cannot be updated at this time. Attempt {0}", counter + 1);
                System.Threading.Thread.Sleep(1000);
                counter++;
            }

            if (ex != null)
            {
                throw ex;
            }

            return true;
        }

        /// <summary>
        /// Get's the base document of the given line number of this document
        /// </summary>
        /// <param name="lineNum">The line number</param>
        /// <returns>The base document, or null if there isn't one</returns>
        public Documents GetLineBaseDocument(int lineNum)
        {
            Documents baseDoc = null;

            var line = this.Document.Lines;

            try
            {
                int currentLine = this.GetLineIndex(line.LineNum);
                line.SetCurrentLine(currentLine);

                try
                {
                    int baseEntry = line.BaseEntry;

                    if (baseEntry != 0)
                    {
                        var baseType = (BoObjectTypes)line.BaseType;
                        int baseLine = line.BaseLine;

                        baseDoc = (Documents)this.Company.GetBusinessObject(baseType);
                        baseDoc.GetByKey(baseEntry);
                    }
                }
                catch (Exception)
                {
                    baseDoc = null;
                }
            }
            finally
            {
                COMHelper.Release(ref line);
            }

            return baseDoc;
        }

        /// <summary>
        /// Gets a dictionary mapping this document's line numbers to the quantities that have been
        /// pulled forward to a target document
        /// </summary>
        /// <returns>The dictionary mapping line numbers to targetted quantities</returns>
        public IDictionary<int, int> GetLineTargetQuantities()
        {
            IDictionary<int, int> lineTargets = new Dictionary<int, int>();

            // Start by assuming nothing has been pulled forward
            var doclines = this.Lines;

            foreach (Document_Lines line in doclines)
            {
                lineTargets[line.LineNum] = 0;
            }

            // Now get all Target documents for each line
            ICollection<int> keys = new List<int>();
            foreach (int lineNum in lineTargets.Keys)
            {
                keys.Add(lineNum);
            }


            foreach (int lineNum in keys)
            {
                IList<Documents> targets = this.GetTargetDocuments(lineNum);
                int targettedQuantity = 0;

                // Go over the document lines, counting the quantities where the base doc/line are this one
                foreach (Documents target in targets)
                {
                    var sapTargetLines = target.Lines;

                    try
                    {
                        sapTargetLines.SetCurrentLine(0);
                        var targetLines = new Lines(sapTargetLines);
                        try
                        {
                            foreach (Document_Lines targetLine in targetLines)
                            {
                                if (targetLine.BaseEntry == this.Document.DocEntry && targetLine.BaseLine == lineNum)
                                {
                                    targettedQuantity += (int)targetLine.Quantity;
                                }
                            }
                        }
                        finally
                        {
                            targetLines.Dispose();
                        }
                    }
                    finally
                    {
                        COMHelper.ReleaseOnly(target);
                    }                    
                }

                targets = null;
                lineTargets[lineNum] = targettedQuantity;
            }

            return lineTargets;
        }

        /// <summary>
        /// Gets A List of all target documents for the given line number of the current document
        /// </summary>
        /// <param name="lineNum">The line number</param>
        /// <returns>A list of all target documents, with their current line set to the corresponding line</returns>
        public IList<Documents> GetTargetDocuments(int lineNum)
        {
            IList<Documents> documents = new List<Documents>();

            string targetTable = SapDocumentTables.GetDocumentRowTable(this.TargetType);

            string query =
                string.Format(
                    @"SELECT DocEntry, LineNum FROM {0} WHERE BaseEntry = {1} AND BaseType = {2} AND BaseLine = {3}",
                    targetTable, 
                    this.Document.DocEntry, 
                    (int)this.DocumentType, 
                    lineNum);

            using (var rows = new RecordsetAdapter(this.Company, query))
            {
                while (!rows.EoF)
               {
                   int docEntry = Convert.ToInt32(rows.FieldValue("DocEntry"));

                   var doc = (Documents)this.Company.GetBusinessObject(this.TargetType);
                   doc.GetByKey(docEntry);

                   documents.Add(doc);

                   rows.MoveNext();
               }
           }

            return documents;
        }

        /// <summary>
        /// Gets all the base documents for this document
        /// </summary>
        /// <returns>A list of all documents that are base documents to this one</returns>
        public IList<Documents> GetAllBaseDocuments()
        {
            IList<Documents> allBases = new List<Documents>();
            IList<int> allBaseDocEntries = new List<int>();
            foreach (Document_Lines line in this.Lines)
            {
                Documents baseDoc = this.GetLineBaseDocument(line.LineNum);
                if (!allBaseDocEntries.Contains(baseDoc.DocEntry))
                {
                    allBaseDocEntries.Add(baseDoc.DocEntry);
                    allBases.Add(baseDoc);
                }
            }

            return allBases;
        }

        /// <summary>
        /// Gets all target documents to this one
        /// </summary>
        /// <returns>A list of all target documents</returns>
        public IList<Documents> GetAllTargetDocuments()
        {
            IList<Documents> allTargets = new List<Documents>();
            IList<int> allTargetDocEntries = new List<int>();
            foreach (Document_Lines line in this.Lines)
            {
                IList<Documents> targets = this.GetTargetDocuments(line.LineNum);

                foreach (Documents target in targets)
                {
                    if (!allTargetDocEntries.Contains(target.DocEntry))
                    {
                        allTargetDocEntries.Add(target.DocEntry);
                        allTargets.Add(target);
                    }
                }
            }

            return allTargets;
        }

        /// <summary>
        /// Gets the line indexes containing the given item code
        /// </summary>
        /// <param name="itemCode">The item code to search for</param>
        /// <returns>All line indexes containing the given item code</returns>
        public IList<int> GetLineNumberForItem(string itemCode)
        {
            IList<int> itemLines = new List<int>();

            Document_Lines docLines = this.Document.Lines;

            try
            {
                for (int i = 0; i < docLines.Count; i++)
                {
                    docLines.SetCurrentLine(i);

                    if (docLines.ItemCode == itemCode)
                    {
                        itemLines.Add(i);
                    }
                }

                docLines.SetCurrentLine(0);
            }
            finally
            {
                COMHelper.Release(ref docLines);
            }

            return itemLines;
        }

        /// <summary>
        /// Gets the line num from document.
        /// </summary>
        /// <param name="baseEntry">The base entry.</param>
        /// <param name="baseLine">The base line.</param>
        /// <returns>The lineNum of the line based on the given document and docLine</returns>
        public int GetLineNumFromDocument(int baseEntry, int baseLine)
        {
            int lineNum = -1;

            using (var docLines = new Lines(this.Document.Lines))
            {
                foreach (Document_Lines line in docLines)
                {
                    if (line.BaseEntry == baseEntry && line.BaseLine == baseLine)
                    {
                        lineNum = line.LineNum;
                        break;
                    }
                }
            }

            return lineNum;
        }

        /// <summary>
        /// Gets the line.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns>The Document Lines</returns>
        public Document_Lines GetLine(int lineNum)
        {
            Document_Lines docLines = this.Document.Lines;

            for (int lineIndex = 0; lineIndex < docLines.Count; lineIndex++)
            {
                docLines.SetCurrentLine(lineIndex);
                if (lineNum == docLines.LineNum)
                {
                    break;
                }
            }

            return docLines;
        }

        /// <summary>
        /// Sets the line.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns>True if the line is found</returns>
        public bool SetLine(int lineNum)
        {
            Document_Lines docLines = this.Document.Lines;

            try
            {
                for (int lineIndex = 0; lineIndex < docLines.Count; lineIndex++)
                {
                    docLines.SetCurrentLine(lineIndex);
                    if (lineNum == docLines.LineNum)
                    {
                        this.Lines.SetCurrentLine(lineIndex);
                        return true;
                    }
                }
            }
            finally
            {
                COMHelper.Release(ref docLines);
            }

            return false;
        }

        /// <summary>
        /// Gets the index of the line by.
        /// </summary>
        /// <param name="lineIndex">Index of the line.</param>
        /// <returns>The Document Lines</returns>
        public Document_Lines GetLineByIndex(int lineIndex)
        {
            Document_Lines docLines = this.Document.Lines;

            try
            {
                docLines.SetCurrentLine(lineIndex);
            }
            finally
            {
                COMHelper.Release(ref docLines);   
            }

            return this.Document.Lines;
        }

        /// <summary>
        /// Gets the index of the line.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns>The index of the line</returns>
        public int GetLineIndex(int lineNum)
        {
            Document_Lines docLines = this.Document.Lines;

            try
            {
                for (int lineIndex = 0; lineIndex < docLines.Count; lineIndex++)
                {
                    docLines.SetCurrentLine(lineIndex);
                    if (lineNum == docLines.LineNum)
                    {
                        return lineIndex;
                    }
                }

                return -1;
            }
            finally
            {
                COMHelper.Release(ref docLines);
            }
        }

        /// <summary>
        /// Sets the index of the line by.
        /// </summary>
        /// <param name="lineIndex">Index of the line.</param>
        /// <returns>True if the line is found</returns>
        public bool SetLineByIndex(int lineIndex)
        {
            Document_Lines docLines = this.Document.Lines;

            try
            {
                docLines.SetCurrentLine(lineIndex);
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                COMHelper.Release(ref docLines);
            }
        }

        /// <summary>
        /// Gets the Closed Quantities on a per-item bases
        /// </summary>
        /// <remarks>
        /// In this case, 'Closed' means that a Target document exists containing the item, 
        /// it does not include 'Cancelled' quantities
        /// </remarks>
        /// <returns>A dictionary mapping item codes to closed quantities</returns>
        public IDictionary<string, int> GetClosedQuantitiesByItem()
        {
            this.Task = string.Format("Getting closed quantities by item for Document {0} of type {1}", this.Document.DocEntry, this.DocumentCode);
            IDictionary<string, int> docLines = new Dictionary<string, int>();

            // Start by assuming all lines are closed
            var documentLines = this.Document.Lines;
            documentLines.SetCurrentLine(0);
            try
            {
                foreach (Document_Lines line in this.Lines)
                {
                    if (!docLines.ContainsKey(line.ItemCode))
                    {
                        docLines[line.ItemCode] = 0;
                    }

                    docLines[line.ItemCode] += (int)line.Quantity;
                }

                IDictionary<string, int> openQuantities = this.GetOpenQuantitiesByItem();

                foreach (string itemCode in openQuantities.Keys)
                {
                    docLines[itemCode] -= openQuantities[itemCode];
                }

                return docLines;
            }
            finally
            {
                COMHelper.Release(ref documentLines);   
            }
        }

        /// <summary>
        /// Gets the Open Quantities on a per-item basis
        /// </summary>
        /// <returns>A dictionary mapping item codes to remaining open quantities</returns>
        public IDictionary<string, int> GetOpenQuantitiesByItem()
        {
            IDictionary<string, int> docLines = new Dictionary<string, int>();

            // Start by assuming all lines are open
            var documentLines = this.Document.Lines;
            documentLines.SetCurrentLine(0);

            try
            {
                foreach (Document_Lines line in this.Lines)
                {
                    if (!docLines.ContainsKey(line.ItemCode))
                    {
                        docLines[line.ItemCode] = 0;
                    }

                    docLines[line.ItemCode] += (int)line.Quantity;
                }

                // Subtract those with targets
                documentLines.SetCurrentLine(0);
                foreach (Document_Lines line in this.Lines)
                {
                    IList<Documents> targetDocuments = this.GetTargetDocuments(line.LineNum);

                    foreach (Documents targetDocument in targetDocuments)
                    {
                        try
                        {
                            var targetDocumentLines = targetDocument.Lines;
                            targetDocumentLines.SetCurrentLine(0);

                            using (var targetDocLines = new Lines(targetDocument.Lines))
                            {
                                foreach (Document_Lines targetLine in targetDocLines)
                                {
                                    if (targetLine.ItemCode.Equals(line.ItemCode))
                                    {
                                        // Subtract the target quantity from the running tally on that item
                                        docLines[line.ItemCode] = docLines[line.ItemCode] - (int)targetLine.Quantity;
                                    }
                                }
                            }
                        }
                        finally
                        {
                            COMHelper.ReleaseOnly(targetDocument);
                        }
                    }
                }

                return docLines;
            }
            finally
            {
                COMHelper.Release(ref documentLines);
            }
        }

        /// <summary>
        /// Gets the quantity of each item code found in the document
        /// </summary>
        /// <returns>A dictionary mapping item codes to quantities</returns>
        public IDictionary<string, int> GetItemQuantities()
        {
            IDictionary<string, int> quantities = new Dictionary<string, int>();

            using (var docLines = new Lines(this.Document.Lines))
            {
                foreach (Document_Lines line in docLines)
                {
                    if (!quantities.ContainsKey(line.ItemCode))
                    {
                        quantities[line.ItemCode] = 0;
                    }

                    quantities[line.ItemCode] += (int)line.Quantity;
                }

                return quantities;
            }
        }

        /// <summary>
        /// Gets the item unit prices.
        /// </summary>
        /// <returns>A dictionary mapping item codes to prices</returns>
        public IDictionary<string, double> GetItemUnitPrices()
        {
            IDictionary<string, double> prices = new Dictionary<string, double>();

            using (var doclines = new Lines(this.Document.Lines))
            {
                foreach (Document_Lines line in doclines)
                {
                    if (!prices.ContainsKey(line.ItemCode))
                    {
                        prices[line.ItemCode] = line.Price;
                    }
                }

                return prices;
            }
        }

        /// <summary>
        /// Creates the report.
        /// </summary>
        /// <returns>The path to the created report as a pdf</returns>
        public virtual string CreateReport()
        {
            // Get the report file
            var systemFolder = new SystemFolder(this.Company);

            string reportsFolder = systemFolder.Reports;

            // Look for the default report file in the reports directory
            var reportsFolderInfo = new DirectoryInfo(reportsFolder);
            FileInfo[] files = reportsFolderInfo.GetFiles(this.ReportName);
            if (files.Length != 1)
            {
                ThreadedAppLog.WriteLine("Found {0} report files for type {1}. Looking for a file named {2}", files.Length, this.DocumentCode, this.ReportName);
                return string.Empty;
            }

            // Make sure the file exists
            FileInfo reportFile = files[0];
            if (!reportFile.Exists)
            {
                ThreadedAppLog.WriteLine("The report file {0} does not exist", this.ReportName);
                return string.Empty;
            }

            // Export the report to the default location
            string exportFileName = this.GetReportExportFileName();
            return this.CreateReport(reportFile.FullName, exportFileName);
        }

        /// <summary>
        /// Creates the report.
        /// </summary>
        /// <param name="reportFile">The report file.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>Returns the report path</returns>
        public virtual string CreateReport(string reportFile, string fileName)
        {
            var reportUser = new ReportUser();
            var helper = new CrystalHelper(reportFile)
            {
                ServerName = this.Company.Server,
                DatabaseName = this.Company.CompanyDB,
                UserId = reportUser.Name,
                Password = reportUser.Password,
                IntegratedSecurity = false
            };

            string outputFile = this.Company.AttachMentPath + fileName;
            
            helper.SetParameter("DocEntry", this.Document.DocEntry);
            helper.Export(outputFile, CrystalHelper.CrystalExportFormat.PortableDocFormat);

            return outputFile;
        }

        /// <summary>
        /// Queues an email to be sent with next batch
        /// </summary>
        /// <param name="subject">The email subject line</param>
        /// <param name="body">The email body</param>
        /// <param name="recipients">The intended email recipients</param>
        public void QueueSendDocumentInEmail(string subject, string body, string recipients)
        {
            try
            {
                string documentType = this.DocumentCode.ToString();
                int docEntry = this.Document.DocEntry;
                const string Status = "Q";

                string selectQuery = string.Format("SELECT 1 FROM EmailNotifications where DocumentType = {0} and DocEntry = {1}", this.DocumentCode, this.Document.DocEntry);

                var selectRecordset = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                try
                {
                    selectRecordset.DoQuery(selectQuery);

                    if (!selectRecordset.EoF)
                    {
                        ThreadedAppLog.WriteLine("Email has already been sent for this document. Not re-queing");
                        return;
                    }
                }
                finally
                {
                    COMHelper.Release(ref selectRecordset);
                }

                string insertQuery =
                    string.Format(
                        "INSERT INTO EmailNotifications (DocumentType, DocEntry, [Status], Subject, Body, Recipients) VALUES ('{0}', {1}, '{2}', '{3}', '{4}', '{5}')",
                        documentType, 
                        docEntry, 
                        Status, 
                        subject, 
                        body, 
                        recipients);

                var recordset = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                try
                {
                    recordset.DoQuery(insertQuery);
                }
                finally
                {
                    COMHelper.Release(ref recordset);
                }
            }
            catch (Exception ex)
            {
                ThreadedAppLog.WriteLine("Unable to insert email notification into queue");
                ThreadedAppLog.WriteErrorLine(ex);
            }
        }

        /// <summary>
        /// Emails a document to a given email address.
        /// THis document will be attached as a PDF
        /// </summary>
        /// <param name="recipients">The people receiving the email</param>
        /// <param name="subject">The email subject line</param>
        /// <param name="body">The email body</param>
        /// <param name="attachments">The attachments.</param>
        public void EmailDocument(string recipients, string subject, string body, params string[] attachments)
        {
            var message = (Messages)this.Company.GetBusinessObject(BoObjectTypes.oMessages);

            message.Subject = subject;
            message.MessageText = body;

            message.Recipients.Add();
            message.Recipients.EmailAddress = recipients;
            message.Recipients.SendEmail = BoYesNoEnum.tYES;
            message.Recipients.SendInternal = BoYesNoEnum.tNO;
            message.Recipients.SendFax = BoYesNoEnum.tNO;
            message.Recipients.SendSMS = BoYesNoEnum.tNO;
            message.Recipients.UserCode = this.Document.CardCode;
            message.Recipients.NameTo = "Listener";
            message.Recipients.UserType = BoMsgRcpTypes.rt_ContactPerson;

            var messageAttachments = message.Attachments;

            foreach (string attachment in attachments)
            {
                messageAttachments.Add();
                messageAttachments.Item(0).FileName = attachment;
            }

            ThreadedAppLog.WriteLine("Sending email containing document to {0}", recipients);

            if (message.Add() != 0)
            {
                throw new SapException(this.Company);
            }
        }

        /// <summary>
        /// Creates a new document for a subset of items/quantities of items in orignal document
        /// </summary>
        /// <param name="itemQuantities">The items and quantities to pull forward</param>
        /// <returns>An unsaved Goods Receipt containing the specified items/quantities</returns>
        public Documents CopyQuantitiesToTarget(IDictionary<string, int> itemQuantities)
        {
            var targetDoc = (Documents)this.Company.GetBusinessObject(this.TargetType);

            IDictionary<int, int> quantitiesToReceivePerLine = new Dictionary<int, int>();
            IDictionary<string, int> quantitiesHandled = new Dictionary<string, int>();

            // Initialize the handled quantities to 0 for each item
            foreach (string itemCode in itemQuantities.Keys)
            {
                quantitiesHandled[itemCode] = 0;
            }

            // Determine the base lines to pull forward
            IDictionary<int, int> targetQuantityByLine = this.GetLineTargetQuantities();

            foreach (Document_Lines line in this.Lines)
            {
                // If the line is closed, we can't copy to target
                if (line.Quantity <= targetQuantityByLine[line.LineNum])
                {
                    continue;
                }

                // Check if this item code is needed to be copied
                if (itemQuantities.ContainsKey(line.ItemCode))
                {
                    // Check how many of these items we still need to copy forward
                    int remainingQuantity = itemQuantities[line.ItemCode] - quantitiesHandled[line.ItemCode];

                    if (remainingQuantity > 0)
                    {
                        int quantityToHandle = Math.Min(remainingQuantity, (int)line.Quantity);

                        quantitiesToReceivePerLine[line.LineNum] = quantityToHandle;
                        quantitiesHandled[line.ItemCode] += quantityToHandle;
                    }
                }
            }

            // Fill out the document level info            
            targetDoc.CardCode = this.Document.CardCode;
            targetDoc.SalesPersonCode = this.Document.SalesPersonCode;
            targetDoc.DocType = BoDocumentTypes.dDocument_Items;
            targetDoc.DocumentSubType = BoDocumentSubType.bod_None;
            targetDoc.Address = this.Document.Address;
            targetDoc.ShipToCode = this.Document.ShipToCode;
            targetDoc.Address2 = this.Document.Address2;
            targetDoc.HandWritten = BoYesNoEnum.tNO;
            targetDoc.JournalMemo = string.Format("{0} - {1}", SapDocumentTables.GetDocumentName(this.DocumentType), this.Document.CardCode);
            targetDoc.PaymentGroupCode = this.Document.PaymentGroupCode;
            targetDoc.DocDueDate = DateTime.Today;

            var targetAdapter = BusinessAdapterFactory.CreateDocumentAdapter(this.Company, targetDoc);

            targetAdapter.SetComments(string.Format("Based on {0} {1}", SapDocumentTables.GetDocumentName(this.DocumentType), this.Document.DocEntry));

            // Transfer all the custom fields value
            this.TransferFieldsToAdapter(targetAdapter);

            // Now fill out the lines in the target
            Document_Lines targetLine = targetDoc.Lines;
            Document_Lines baseLine = this.Document.Lines;
            try
            {
                int targetLineNum = 0;
                foreach (int lineNum in quantitiesToReceivePerLine.Keys)
                {
                    if (targetLineNum > 0)
                    {
                        targetLine.Add();
                    }
                    baseLine.SetCurrentLine(lineNum);
                    int lineQty = quantitiesToReceivePerLine[lineNum];

                    targetLine.BaseLine = lineNum;
                    targetLine.BaseEntry = this.Document.DocEntry;
                    targetLine.BaseType = (int)this.Document.DocObjectCode;
                    targetLine.Quantity = lineQty;
                    targetLine.ItemCode = baseLine.ItemCode;
                    targetLine.TaxCode = baseLine.TaxCode;
                    targetLine.Rate = baseLine.Rate;

                    // Transfer the line level UDFs
                    this.TransferLineFieldsToAdapter(targetAdapter);
                    this.TransferLineSerialNumbersToAdapter(targetAdapter);
                    this.TransferLineExpensesToAdapter(targetAdapter);

                    targetLineNum++;
                }

                targetLine.SetCurrentLine(0);
            }
            finally
            {
                COMHelper.Release(ref baseLine);
                COMHelper.Release(ref targetLine);
            }

            return targetDoc;
        }

        /// <summary>
        /// Gets the base adapter.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns>An adapter for the base document of this line</returns>
        public DocumentAdapter GetBaseAdapter(int lineNum)
        {
            int lineIndex = this.GetLineIndex(lineNum);

            DocumentAdapter baseAdapter = null;

            if (this.BaseAdapters == null)
            {
                this.BaseAdapters = new Dictionary<int, DocumentAdapter>();
            }

            if (!this.BaseAdapters.ContainsKey(lineIndex))
            {
                Document_Lines line = this.Document.Lines;
                line.SetCurrentLine(lineIndex);

                if (line.BaseEntry > 0)
                {
                    this.BaseAdapters.Add(lineIndex, BusinessAdapterFactory.CreateDocumentAdapter(this.Company, line.BaseType, line.BaseEntry));
                    baseAdapter = this.BaseAdapters[lineIndex];
                }
            }

            return baseAdapter;
        }

        /// <summary>
        /// Gets the line serial numbers.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns>Serial Numbers for the given line</returns>
        public SerialNumbers GetLineSerialNumbers(int lineNum)
        {
            Document_Lines lines = this.Document.Lines;
            lines.SetCurrentLine(this.GetLineIndex(lineNum));

            SerialNumbers sn = lines.SerialNumbers;

            // If there's no serial number on the current line (and a base document exists), get the SN from the base line
            if (sn.SystemSerialNumber <= 0 && this.Lines.Current.BaseEntry > 0)
            {
                var adapter = this.GetBaseAdapter(lines.LineNum);
                sn = adapter.GetLineSerialNumbers(lines.BaseLine);
            }

            return sn;
        }            

        /// <summary>
        /// Checks if the target document will affect inventory.
        /// </summary>
        /// <param name="targetType">Type of the target document.</param>
        /// <returns><c>true</c> if the target document will affect inventory, <c>false</c> otherwise</returns>
        public bool DoesTargetAffectInventory(BoObjectTypes targetType)
        {
            // Note: This logic is hard-coded, as there is no way to generically determine if a target document will affect inventory. It is necessary to know if the target will affect inventory, as SNs or Batches need to be copied in this case
            bool affectsInventory = false;

            switch (this.DocumentType)
            {
                case BoObjectTypes.oOrders:
                case BoObjectTypes.oPurchaseOrders:
                    affectsInventory = true;
                    break;

                case BoObjectTypes.oDeliveryNotes:
                    switch (targetType)
                    {
                        case BoObjectTypes.oReturns:
                            affectsInventory = true;
                            break;
                    }
                    break;

                case BoObjectTypes.oPurchaseDeliveryNotes:
                    switch (targetType)
                    {
                        case BoObjectTypes.oPurchaseReturns:
                            affectsInventory = true;
                            break;
                    }
                    break;

                case BoObjectTypes.oInvoices:
                case BoObjectTypes.oPurchaseInvoices:
                    if (this.Document.Reserve == BoYesNoEnum.tNO)
                        affectsInventory = true;
                    break;
            }

            return affectsInventory;
        }

        /// <summary>
        /// Transfers the line serial numbers to adapter.
        /// </summary>
        /// <param name="targetAdapter">The target adapter.</param>
        public void TransferLineSerialNumbersToAdapter(DocumentAdapter targetAdapter)
        {
            // Transfer the line level serial numbers
            var masterLines = this.Document.Lines;
            var targetLines = targetAdapter.Document.Lines;

            try
            {
                if (!this.DoesTargetAffectInventory(targetAdapter.DocumentType))
                {
                    // We don't need to transfer serial nubmers if the target does not touch inventory
                    return;
                }

                using (var itemAdapter = new ItemAdapter(this.Company, masterLines.ItemCode))
                {
                    if (!itemAdapter.IsItemSerialized())
                    {
                        // We don't need to transfer serial numbers if the item isn't serializable
                        return;
                    }
                }

                var masterLineSerials = this.GetLineSerialNumbers(masterLines.LineNum);
                var newLineSerials = targetLines.SerialNumbers;

                try
                {
                    // Check if this document has serials, otherwise do a deep search for them
                    if (string.IsNullOrEmpty(masterLineSerials.BatchID))
                    {

                    }

                    if (masterLineSerials.Count > 0)
                    {
                        int masterLineCount = masterLineSerials.Count;
                        for (int j = 0; j < masterLineCount; j++)
                        {
                            masterLineSerials.SetCurrentLine(j);

                            newLineSerials.BaseLineNumber = masterLineSerials.BaseLineNumber;
                            newLineSerials.BatchID = masterLineSerials.BatchID;
                            newLineSerials.ExpiryDate = masterLineSerials.ExpiryDate;
                            newLineSerials.InternalSerialNumber = masterLineSerials.InternalSerialNumber;
                            newLineSerials.Location = masterLineSerials.Location;
                            newLineSerials.ManufactureDate = masterLineSerials.ManufactureDate;
                            newLineSerials.ManufacturerSerialNumber = masterLineSerials.ManufacturerSerialNumber;
                            newLineSerials.Notes = masterLineSerials.Notes;
                            newLineSerials.ReceptionDate = masterLineSerials.ReceptionDate;
                            newLineSerials.SystemSerialNumber = masterLineSerials.SystemSerialNumber;
                            newLineSerials.WarrantyEnd = masterLineSerials.WarrantyEnd;
                            newLineSerials.WarrantyStart = masterLineSerials.WarrantyStart;

                            if (j == (masterLineCount - 1))
                            {
                                newLineSerials.Add();
                            }
                        }
                    }
                }
                finally
                {
                    COMHelper.Release(ref masterLineSerials);
                    COMHelper.Release(ref newLineSerials);
                }
            }
            finally
            {
                COMHelper.Release(ref masterLines);
                COMHelper.Release(ref targetLines);
            }

        }

        /// <summary>
        /// Transfers all Userdefined fields to adapter.
        /// </summary>
        /// <param name="targetAdapter">The target adapter.</param>
        public void TransferFieldsToAdapter(DocumentAdapter targetAdapter)
        {
            var documentUserFields = this.Document.UserFields;
            try
            {
                var documentUserFieldsFields = documentUserFields.Fields;

                try
                {
                    // Transfer all the custom fields value
                    for (int fieldIndex = 0; fieldIndex < documentUserFieldsFields.Count; fieldIndex++)
                    {
                        targetAdapter.SetUserDefinedField(fieldIndex, this.GetUserDefinedField(fieldIndex));
                    }
                }
                finally
                {
                    COMHelper.Release(ref documentUserFieldsFields);
                    documentUserFieldsFields = null;
                }
            }
            finally
            {
                COMHelper.Release(ref documentUserFields);
                documentUserFields = null;
            }
        }

        /// <summary>
        /// Transfers all Userdefined line fields to adapter.
        /// </summary>
        /// <param name="targetAdapter">The target adapter.</param>
        public void TransferLineToAdapter(DocumentAdapter targetAdapter)
        {
            var masterLines = this.Document.Lines;
            var newLines = targetAdapter.Document.Lines;

            try
            {
                if (string.IsNullOrEmpty(masterLines.ItemCode))
                {
                    return;
                }

                if (!string.IsNullOrEmpty(newLines.ItemCode))
                {
                    // This means that the cursor is NOT on a new line
                    newLines.Add();
                }

                newLines.ItemCode = masterLines.ItemCode;
                newLines.Quantity = masterLines.Quantity;
                newLines.Rate = masterLines.Rate;
                newLines.TaxCode = masterLines.TaxCode;

                // Only set base type if the base exists
                if (this.Document.DocEntry > 0)
                {
                    newLines.BaseType = (int)this.DocumentType;
                    newLines.BaseEntry = this.Document.DocEntry;
                    newLines.BaseLine = masterLines.LineNum;
                }

                // Transfer the line level UDFs
                this.TransferLineFieldsToAdapter(targetAdapter);

                // Transfer the line level expenses
                this.TransferLineExpensesToAdapter(targetAdapter);

                // Transfer the Serial Numbers
                this.TransferLineSerialNumbersToAdapter(targetAdapter);
            }
            finally
            {
                COMHelper.Release(ref newLines);
                COMHelper.Release(ref masterLines);
            }
        }

        /// <summary>
        /// Transfers all Userdefined line fields to adapter.
        /// </summary>
        /// <param name="targetAdapter">The target adapter.</param>
        public void TransferLineFieldsToAdapter(DocumentAdapter targetAdapter)
        {
            var baseLine = this.Document.Lines;

            try
            {
                var baseLineUserFields = baseLine.UserFields;
                try
                {
                    var baseLineUserFieldsFields = baseLineUserFields.Fields;

                    try
                    {
                        // Transfer the line level UDFs
                        for (int j = 0; j < baseLineUserFieldsFields.Count; j++)
                        {
                            targetAdapter.SetLineUserDefinedField(j, this.GetLineUserDefinedField(j));
                        }
                    }
                    finally
                    {
                        COMHelper.Release(ref baseLineUserFieldsFields);
                    }
                }
                finally
                {
                    COMHelper.Release(ref baseLineUserFields);
                }
            }
            finally
            {
                COMHelper.Release(ref baseLine);
            }
        }

        /// <summary>
        /// Transfers all Userdefined line fields to adapter.
        /// </summary>
        /// <param name="targetAdapter">The target adapter.</param>
        public void TransferLineExpensesToAdapter(DocumentAdapter targetAdapter)
        {
            // Transfer the line level expenses
            var masterLines = this.Document.Lines;
            var targetLines = targetAdapter.Document.Lines;

            try
            {
                var masterLineExpenses = masterLines.Expenses;
                var newLineExpenses = targetLines.Expenses;

                try
                {
                    if (masterLineExpenses.ExpenseCode != 0)
                    {
                        int masterLineExpenseCount = masterLineExpenses.Count;
                        for (int j = 0; j < masterLineExpenseCount; j++)
                        {
                            masterLineExpenses.SetCurrentLine(j);

                            newLineExpenses.ExpenseCode = masterLineExpenses.ExpenseCode;
                            newLineExpenses.LineTotal = masterLineExpenses.LineTotal;
                            newLineExpenses.GroupCode = masterLineExpenses.GroupCode;
                            newLineExpenses.BaseGroup = masterLineExpenses.GroupCode;

                            if (j == (masterLineExpenseCount - 1))
                            {
                                newLineExpenses.Add();
                            }
                        }
                    }
                }
                finally
                {
                    COMHelper.Release(ref masterLineExpenses);
                    COMHelper.Release(ref newLineExpenses);
                }
            }
            finally
            {
                COMHelper.Release(ref masterLines);
                COMHelper.Release(ref targetLines);
            }
        }

        /// <summary>
        /// Transfers all Userdefined line fields to adapter.
        /// </summary>
        /// <param name="targetAdapter">The target adapter.</param>
        public void TransferExpensesToAdapter(DocumentAdapter targetAdapter)
        {
            DocumentsAdditionalExpenses masterExpenses = this.Document.Expenses;
            DocumentsAdditionalExpenses targetExpenses = targetAdapter.Document.Expenses;

            try
            {
                int expenseLine = 0;
                while (expenseLine < masterExpenses.Count)
                {
                    masterExpenses.SetCurrentLine(expenseLine);
                    if (masterExpenses.ExpenseCode != 0)
                    {
                        targetExpenses.ExpenseCode = masterExpenses.ExpenseCode;
                        targetExpenses.TaxCode = masterExpenses.TaxCode;
                        targetExpenses.VatGroup = masterExpenses.VatGroup;
                        targetExpenses.LineTotal = masterExpenses.LineTotal;
                        if (expenseLine < masterExpenses.Count - 1)
                        {
                            targetExpenses.Add();
                        }
                    }

                    expenseLine++;
                }
            }
            finally
            {
                COMHelper.Release(ref targetExpenses);
                COMHelper.Release(ref masterExpenses);
            }
        }

        /// <summary>
        /// Sets the user defined field.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <param name="fieldValue">The field value.</param>
        public void SetUserDefinedField(object fieldIndex, object fieldValue)
        {
            COMHelper.UserDefinedFieldValue(this.Document.UserFields, fieldIndex, fieldValue);
        }

        /// <summary>
        /// Sets the user defined field on the current line.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <param name="fieldValue">The field value.</param>
        public void SetLineUserDefinedField(object fieldIndex, object fieldValue)
        {
            Document_Lines line = this.Document.Lines;
            try
            {
                UserFields uf = line.UserFields;
                try
                {
                    Fields fields = uf.Fields;
                    try
                    {
                        Field field = fields.Item(fieldIndex);
                        try
                        {
                            field.Value = fieldValue;
                        }
                        finally
                        {
                            COMHelper.Release(ref field);
                        }
                    }
                    finally
                    {
                        COMHelper.Release(ref fields);
                    }
                }
                finally
                {
                    COMHelper.Release(ref uf);
                }
            }
            finally
            {
                COMHelper.Release(ref line);
            }
        }

        /// <summary>
        /// Gets the user defined field from the Document.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <returns>The value of the field</returns>
        public object GetUserDefinedField(object fieldIndex)
        {
            return COMHelper.UserDefinedFieldValue(this.Document.UserFields, fieldIndex);
        }

        /// <summary>
        /// Gets the user defined field from the current line.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <returns>The value of the field</returns>
        public object GetLineUserDefinedField(object fieldIndex)
        {
            Document_Lines line = this.Document.Lines;
            try
            {
                UserFields uf = line.UserFields;
                try
                {
                    Fields fields = uf.Fields;
                    try
                    {
                        Field field = fields.Item(fieldIndex);
                        try
                        {
                            return field.Value;
                        }
                        finally
                        {
                            COMHelper.Release(ref field);
                        }
                    }
                    finally
                    {
                        COMHelper.Release(ref fields);
                    }
                }
                finally
                {
                    COMHelper.Release(ref uf);
                }
            }
            finally
            {
                COMHelper.Release(ref line);
            }
        }

        /// <summary>
        /// Gets the contact.
        /// </summary>
        /// <returns>The contact employee object</returns>
        public ContactEmployees GetContact()
        {
            var bp = (SAPbobsCOM.BusinessPartners)this.Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);

            try
            {
                // Try to find the Contact Employee referenced by the document
                int employeeCode = this.Document.ContactPersonCode;
                bool foundEmployee = false;
                bp.GetByKey(this.Document.CardCode);

                var employee = bp.ContactEmployees;

                int employeeCount = employee.Count;
                for (int i = 0; i < employeeCount; i++)
                {
                    employee.SetCurrentLine(i);

                    if (employee.InternalCode == employeeCode)
                    {
                        foundEmployee = true;
                        break;
                    }
                }

                // Check if we found the employee
                if (!foundEmployee)
                {
                    ThreadedAppLog.WriteLine("Unable to find contact person with code {0} on business partner {1}", employeeCode, this.Document.CardCode);
                    employee = null;
                }

                return employee;
            }
            finally
            {
                COMHelper.Release(ref bp);
            }
        }

        /// <summary>
        /// Copies all to draft.
        /// </summary>
        /// <returns>The Draft document created.</returns>
        public DraftAdapter CopyAllToDraft()
        {
            var draft = this.CopyAllToDocument(BoObjectTypes.oDrafts);
            var draftAdapter = new DraftAdapter(this.Company, draft);
            draft.DocObjectCode = this.DocumentType;
            draft.DocDate = this.Document.DocDate;
            draft.DocDueDate = this.Document.DocDueDate;
            draft.TaxDate = this.Document.TaxDate;

            using (var draftLines = new Lines(draft.Lines))
            {
                using (var thisLines = new Lines(this.Document.Lines))
                {
                    foreach (Document_Lines draftLine in draftLines)
                    {
                        Document_Lines thisLine = thisLines.Current;

                        if (thisLine.BaseType != -1)
                        {
                            draftLine.BaseLine = thisLine.BaseLine;
                            draftLine.BaseEntry = thisLine.BaseEntry;
                            draftLine.BaseType = thisLine.BaseType;
                        }

                        thisLines.MoveNext();
                    }
                }
            }

            return draftAdapter;
        }

        /// <summary>
        /// Copies the lines to draft.
        /// </summary>
        /// <param name="linesToCopy">The lines to copy.</param>
        /// <returns>The Draft document created.</returns>
        public Documents CopyLinesToDraft(int[] linesToCopy)
        {
            Documents draft = this.CopyLinesToDocument(BoObjectTypes.oDrafts, linesToCopy);
            return draft;
        }

        /// <summary>
        /// Adds the lines to document.
        /// </summary>
        /// <param name="newDocument">The new document.</param>
        /// <param name="docLines">The doc lines.</param>
        public void AddLinesToDocument(Documents newDocument, params int[] docLines)
        {
            // Transfer each line item over
            Document_Lines newLines = newDocument.Lines;
            Document_Lines masterLines = this.Document.Lines;

            try
            {
                foreach (int line in docLines)
                {
                    masterLines.SetCurrentLine(line);

                    newLines.Add();

                    newLines.ItemCode = masterLines.ItemCode;
                    newLines.Quantity = masterLines.Quantity;
                    newLines.Rate = masterLines.Rate;
                    newLines.BaseType = (int)SapDocumentTables.GetDocumentAPARDocumentType(this.DocumentType);
                    newLines.BaseEntry = this.Document.DocEntry;
                    newLines.BaseLine = masterLines.LineNum;
                }
            }
            finally
            {
                COMHelper.Release(ref newLines);
                COMHelper.Release(ref masterLines);
            }
        }

        /// <summary>
        /// Deletes the draft if the current document was generated from a draft.
        /// </summary>
        public void DeleteDraft()
        {
            int deleteDraftId = this.DraftId;

            this.Task = string.Format("Deleting Draft for document [{0}: {1}] - Draft: {2}", this.DocumentCode, this.Document.DocEntry, deleteDraftId);

            if (deleteDraftId > 0)
            {
                var draft = (Documents)this.Company.GetBusinessObject(BoObjectTypes.oDrafts);

                try
                {
                    if (draft.GetByKey(deleteDraftId))
                    {
                        draft.Remove();
                    }
                }
                finally
                {
                    COMHelper.Release(ref draft);   
                }
            }
        }

        /// <summary>
        /// Determines whether [is document from draft].
        /// </summary>
        /// <returns>
        /// <c>true</c> if [is document from draft]; otherwise, <c>false</c>.
        /// </returns>
        public bool IsDocumentFromDraft()
        {
            return this.DraftId > 0;
        }

        /// <summary>
        /// Validates the participant address.
        /// </summary>
        /// <returns>The validation result (empty is succesfull.</returns>
        public string ValidateParticipantAddress()
        {
            string errors = string.Empty;

            if (string.IsNullOrEmpty(this.GetUserDefinedField("U_XX_ShipToName").ToString()))
            {
                errors += "Participant ship to name is missing.\n";
            }

            if (string.IsNullOrEmpty(this.GetUserDefinedField("U_XX_PAXFirstName").ToString()))
            {
                errors += "Participant first name is missing.\n";
            }

            if (string.IsNullOrEmpty(this.GetUserDefinedField("U_XX_PAXLastName").ToString()))
            {
                errors += "Participant last name is missing.\n";
            }

            if (string.IsNullOrEmpty(this.GetUserDefinedField("U_XX_ShipToAddress1").ToString()))
            {
                errors += "Participant address is missing.\n";
            }

            if (string.IsNullOrEmpty(this.GetUserDefinedField("U_XX_ShipToCity").ToString()))
            {
                errors += "Participant address (city) is missing.\n";
            }

            if (string.IsNullOrEmpty(this.GetUserDefinedField("U_XX_ShipToProvince").ToString()))
            {
                errors += "Participant address (province) is missing.\n";
            }

            if (string.IsNullOrEmpty(this.GetUserDefinedField("U_XX_ShipToCountry").ToString()))
            {
                errors += "Participant address (country) is missing.\n";
            }
            
            return errors;
        }

        /// <summary>
        /// The send financial notification.
        /// </summary>
        /// <param name="subject">The subject of the notification.</param>
        /// <param name="body">The body of the notification.</param>
        public void NotifyDocumentCreator(string subject, string body)
        {
            var sql = new SqlHelper();
            sql.Builder.AppendLine("SELECT T0.User_Code AS UserCode ");
            sql.Builder.AppendFormat("FROM OUSR T0 JOIN [{0}] T1 ON T1.UserSign = T0.INTERNAL_K ", this.DocumentTable);
            sql.Builder.AppendFormat("WHERE ISNULL(E_mail, '') <> '' AND T1.DocEntry = {0}", this.Document.DocEntry);

            using (var rs = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                if (!rs.EoF)
                {
                    // Send Email
                    var message = (Messages)Company.GetBusinessObject(BoObjectTypes.oMessages);

                    try
                    {
                        message.Subject = subject;
                        message.MessageText = body;
                        string recipientUserCode = rs.FieldValue("UserCode").ToString();

                        var messageRecipients = message.Recipients;

                        try
                        {
                            if (!string.IsNullOrEmpty(recipientUserCode))
                            {
                                message.Recipients.UserCode = recipientUserCode;
                                message.Recipients.SendInternal = BoYesNoEnum.tYES;
                            }
                        }
                        finally
                        {
                            COMHelper.Release(ref messageRecipients);
                        }

                        message.Priority = BoMsgPriorities.pr_High;

                        if (message.Add() != 0)
                        {
                            throw new SapException(this.Company);
                        }
                    }
                    finally
                    {
                        COMHelper.Release(ref message);   
                    }
                }
            }
        }

        /// <summary>
        /// Checks if the document exists.
        /// </summary>
        /// <returns>True if the document exists. False otherwise</returns>
        public bool Exists()
        {
            bool exists = false;

            if (this.Document != null)
            {
                using (var rs = new RecordsetAdapter(this.Company, string.Format("SELECT TOP 1 DocEntry FROM {0} WHERE DOCENTRY = {1}", this.DocumentTable, this.Document.DocEntry)))
                {
                    if (!rs.EoF)
                    {
                        exists = true;
                    }
                }
            }

            return exists;
        }

        /// <summary>
        /// Sets the comments.
        /// </summary>
        /// <param name="comments">The comments.</param>
        public void SetComments(string comments)
        {
            this.Document.Comments = DocumentMaxLengths.TrimLength(comments, DocumentMaxLengths.Comments);
        }

        /// <summary>
        /// Appends the comments.
        /// </summary>
        /// <param name="newComments">The new comments.</param>
        public void AppendComments(string newComments)
        {
            this.SetComments(this.Document.Comments + "\n" + newComments);
        }

        /// <summary>
        /// Releases the Documents COM Object.
        /// </summary>
        protected override void Release()
        {
            if (this._businessPartner != null)
            {
                this._businessPartner.Dispose();
            }

            if (this._lines != null)
            {
                this._lines.Dispose();
                this._lines = null;
            }

            if (this.BaseAdapter != null)
            {
                this.BaseAdapter.Dispose();
                this.BaseAdapter = null;
            }

            if (this.Document != null)
            {
                COMHelper.ReleaseOnly(this.Document);
                this.Document = null;
            }            
        }

        /// <summary>
        /// Gets all item codes that are in the given group
        /// </summary>
        /// <param name="groupId">The id of the item group to retrieve</param>
        /// <returns>A list of ItemCodes belonging to the given group, and found in this document</returns>
        protected IList<string> GetItemCodesInGroup(int groupId)
        {
            IList<string> itemCodes = new List<string>();

            // Build the query
            string query =
                string.Format(
                    @"SELECT row.ItemCode from {0} row INNER JOIN OITM oitm ON oitm.ItemCode = row.ItemCode WHERE row.DocEntry = {1} AND oitm.ItmsGrpCod = '{2}'",
                    this.DocumentRowTable, 
                    this.Document.DocEntry, 
                    groupId);

            // Execute the query
            using (var recordset = new RecordsetAdapter(this.Company, query))
            {
                recordset.MoveFirst();
                while (!recordset.EoF)
                {
                    itemCodes.Add(recordset.FieldValue("ItemCode").ToString());
                    recordset.MoveNext();
                }

                return itemCodes;
            }
        }

        /// <summary>
        /// Creates a new document, based on this one
        /// </summary>
        /// <remarks>
        /// This does NOT add the newly created document to the database. It is the 
        /// responsibility of the calling object to call the Add() method on the document, 
        /// after setting any additional fields. 
        /// </remarks>
        /// <param name="documentType">The type of document to create</param>
        /// <returns>The newly created (unsaved) document</returns>
        protected Documents CopyAllToDocument(BoObjectTypes documentType)
        {
            var newDocument = (Documents)this.Company.GetBusinessObject(documentType);

            newDocument.CardCode = this.Document.CardCode;
            newDocument.SalesPersonCode = this.Document.SalesPersonCode;
            newDocument.DocType = BoDocumentTypes.dDocument_Items;
            newDocument.DocumentSubType = BoDocumentSubType.bod_None;
            newDocument.Address = this.Document.Address;
            newDocument.ShipToCode = this.Document.ShipToCode;
            newDocument.Address2 = this.Document.Address2;
            newDocument.HandWritten = BoYesNoEnum.tNO;
            newDocument.JournalMemo = string.Format("{0} - {1}", SapDocumentTables.GetDocumentName(documentType), this.Document.CardCode);                        
            newDocument.Comments = DocumentMaxLengths.TrimLength(string.Format("Based on {0} {1}\n{2}", SapDocumentTables.GetDocumentName(this.DocumentType), this.Document.DocEntry, this.Document.Comments), DocumentMaxLengths.Comments);
            newDocument.PaymentGroupCode = this.Document.PaymentGroupCode;
            newDocument.DocDueDate = DateTime.Today;

            string bpDefaultPaymentMethod = this.BusinessPartner.DefaultPaymentMethod;
            newDocument.PaymentMethod = bpDefaultPaymentMethod == string.Empty
                                                ? this.Document.PaymentMethod
                                                : bpDefaultPaymentMethod;

            DocumentAdapter newDocumentAdapter = BusinessAdapterFactory.CreateDocumentAdapter(this.Company, newDocument);

            // Transfer all the custom fields value
            this.TransferFieldsToAdapter(newDocumentAdapter);

            // Transfer each line item over
            var newLines = newDocument.Lines;
            var masterLines = this.Document.Lines;

            try
            {
                // Loop through all the lines in the master document
                for (int i = 0; i < masterLines.Count; i++)
                {
                    masterLines.SetCurrentLine(i);
                    this.TransferLineToAdapter(newDocumentAdapter);
                }

                // Move the cursor back to the first line
                newLines.SetCurrentLine(0);

                // Copy the expenses over
                this.TransferExpensesToAdapter(newDocumentAdapter);
            }
            finally
            {
                COMHelper.Release(ref newLines);
                COMHelper.Release(ref masterLines);
            }

            return newDocument;
        }

        /// <summary>
        /// Loads a document from the SAP database
        /// </summary>
        /// <param name="docEntry">The document's ID (docEntry)</param>
        /// <returns>The document object</returns>
        protected virtual Documents LoadDocument(int docEntry)
        {
            // Pre-condition
            if (this.Company == null)
            {
                throw new NullReferenceException("The company for the DocumentAdapter is not set.");
            }

            var document = (Documents)this.Company.GetBusinessObject(this.DocumentType);
            document.GetByKey(docEntry);

            return document;
        }

        /// <summary>
        /// Creates a new document, based on this one
        /// </summary>
        /// <param name="documentType">The type of document to create</param>
        /// <param name="docLines">The doc lines. This expects an array of line indexes, not the LineNum values</param>
        /// <returns>The newly created (unsaved) document</returns>
        /// <remarks>
        /// This does NOT add the newly created document to the database. It is the
        /// responsibility of the calling object to call the Add() method on the document,
        /// after setting any additional fields.
        /// </remarks>
        protected Documents CopyLinesToDocument(BoObjectTypes documentType, params int[] docLines)
        {
            var newDocument = (Documents)this.Company.GetBusinessObject(documentType);
            newDocument.CardCode = this.Document.CardCode;
            newDocument.SalesPersonCode = this.Document.SalesPersonCode;
            newDocument.DocType = BoDocumentTypes.dDocument_Items;
            newDocument.DocumentSubType = BoDocumentSubType.bod_None;
            newDocument.Address = this.Document.Address;
            newDocument.ShipToCode = this.Document.ShipToCode;
            newDocument.Address2 = this.Document.Address2;
            newDocument.HandWritten = BoYesNoEnum.tNO;
            newDocument.JournalMemo = string.Format("{0} - {1}", SapDocumentTables.GetDocumentName(documentType), this.Document.CardCode);
            newDocument.Comments = DocumentMaxLengths.TrimLength(string.Format("Based on {0} {1}", SapDocumentTables.GetDocumentName(this.DocumentType), this.Document.DocEntry), DocumentMaxLengths.Comments);
            newDocument.PaymentGroupCode = this.Document.PaymentGroupCode;
            newDocument.DocDueDate = DateTime.Today;

            DocumentAdapter newDocumentAdapter = BusinessAdapterFactory.CreateDocumentAdapter(this.Company, newDocument);

            // Transfer all the custom fields value
            this.TransferFieldsToAdapter(newDocumentAdapter);

            // Transfer each line item over
            var newLines = newDocument.Lines;
            var masterLines = this.Document.Lines;

            try
            {
                foreach (int line in docLines)
                {
                    masterLines.SetCurrentLine(line);
                    this.TransferLineToAdapter(newDocumentAdapter);
                }
            }
            finally
            {
                COMHelper.Release(ref newLines);
                COMHelper.Release(ref masterLines);
            }

            return newDocument;
        }

        /// <summary>
        /// Gets the name of the report export file.
        /// </summary>
        /// <returns>The name of the report export file</returns>
        protected string GetReportExportFileName()
        {
            return SapDocumentTables.GetReportName(this.DocumentType) + "_" + this.Document.DocEntry + "_" + DateTime.Now.ToString("ddMMyyyy") + ".pdf";
        }
        #endregion Methods
    }
}
