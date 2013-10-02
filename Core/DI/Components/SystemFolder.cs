// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SystemFolder.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Expose the system Folders
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.Components
{
    /// <summary>
    /// Expose the system Folders
    /// </summary>
    public class SystemFolder
    {
        /// <summary>
        /// The company
        /// </summary>
        private readonly SAPbobsCOM.Company _company;

        /// <summary>
        /// Initializes a new instance of the <see cref="SystemFolder"/> class.
        /// </summary>
        /// <param name="instanceCompany">The instance company.</param>
        public SystemFolder(SAPbobsCOM.Company instanceCompany)
        {
            this._company = instanceCompany;
        }

        /// <summary>
        /// Gets the reports folder.
        /// </summary>
        /// <value>The reports folder.</value>
        public string Reports
        {
            get
            {
                return ValidatePath(this._company.BitMapPath).Replace("Bitmaps\\", "Reports\\");
            }
        }

        /// <summary>
        /// Gets the patches.
        /// </summary>
        /// <value>The patches.</value>
        public string Patches
        {
            get
            {
                return ValidatePath(this._company.BitMapPath).Replace("Bitmaps\\", "Patches\\");
            }
        }


        /// <summary>
        /// Gets the patches.
        /// </summary>
        /// <value>The patches.</value>
        public string UserLogs
        {
            get
            {
                return ValidatePath(this._company.BitMapPath).Replace("Bitmaps\\", "UserLogs\\" + this._company.UserName + "\\");
            }
        }

        /// <summary>
        /// Gets the Images folder.
        /// </summary>
        /// <value>The Images folder.</value>
        public string Images
        {
            get
            {
                return ValidatePath(this._company.BitMapPath);
            }
        }

        /// <summary>
        /// Gets the Attachments folder.
        /// </summary>
        /// <value>The Attachments folder.</value>
        public string Attachments
        {
            get
            {
                return ValidatePath(this._company.AttachMentPath);
            }
        }

        /// <summary>
        /// Gets the ExcelDocuments folder.
        /// </summary>
        /// <value>The ExcelDocuments folder.</value>
        public string ExcelDocuments
        {
            get
            {
                return ValidatePath(this._company.ExcelDocsPath);
            }
        }

        /// <summary>
        /// Gets the WordDocuments folder.
        /// </summary>
        /// <value>The WordDocuments folder.</value>
        public string WordDocuments
        {
            get
            {
                return ValidatePath(this._company.WordDocsPath);
            }
        }

        /// <summary>
        /// Gets the DisplayListReports folder.
        /// </summary>
        /// <value>The DisplayListReports folder.</value>
        public string DisplayListReports
        {
            get
            {
                return this.Reports + "DisplayList\\";
            }
        }

        /// <summary>
        /// Gets the application server folder.
        /// </summary>
        /// <value>Returns the application server folder.</value>
        public string ApplicationServer
        {
            get
            {
                return ValidatePath(this._company.BitMapPath).Replace("B1_Data\\Bitmaps\\", string.Empty);
            }
        }

        /// <summary>
        /// Gets the log folder.
        /// </summary>
        /// <value>Returns the log folder.</value>
        public string LogFolder
        {
            get
            {
                return this.ExcelDocuments.Replace("ExclDocs\\", "") + "Log\\";
            }
        }

        /// <summary>
        /// Gets the temporary folder.
        /// </summary>
        /// <value>The temporary folder.</value>
        public string TemporaryFolder
        {
            get
            {
                return this.ApplicationServer + "Temp\\";
            }
        }

        /// <summary>
        /// Gets the listener folder.
        /// </summary>
        /// <value>The listener folder.</value>
        public string ListenerFolder
        {
            get
            {
                return this.ApplicationServer + "Listeners\\";
            }
        }

        /// <summary>
        /// Gets the service folder.
        /// </summary>
        /// <value>The service folder.</value>
        public string ServiceFolder
        {
            get
            {
                return this.ApplicationServer + "Service\\";
            }
        }

        /// <summary>
        /// Validates the path.
        /// </summary>
        /// <param name="inputPath">The input path.</param>
        /// <returns>Teh path with \ at the end</returns>
        private static string ValidatePath(string inputPath)
        {
            if (!string.IsNullOrEmpty(inputPath))
            {
                if (!inputPath.EndsWith("\\"))
                {
                    inputPath += "\\";
                }
            }

            return inputPath;
        }
    }
}
