//-----------------------------------------------------------------------
// <copyright file="FileSelectorHelper.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.UI.Helpers
{
    #region Using
    using System;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    #endregion

    /// <summary>
    /// Instance of the FileSelector class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public class FileSelector
    {
        #region Fields
        /// <summary>
        /// The file dialog box
        /// </summary>
        private readonly OpenFileDialog fileDialog;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="FileSelector"/> class.
        /// </summary>
        public FileSelector()
        {
            this.fileDialog = new OpenFileDialog();
        }
        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets the name of the file.
        /// </summary>
        /// <value>The name of the file.</value>
        public string FileName
        {
            get { return this.fileDialog.FileName; }

            set { this.fileDialog.FileName = value; }
        }

        /// <summary>
        /// Gets or sets the filter.
        /// </summary>
        /// <value>The file filter.</value>
        public string Filter
        {
            get { return this.fileDialog.Filter; }
            set { this.fileDialog.Filter = value; }
        }

        /// <summary>
        /// Gets or sets the initial directory.
        /// </summary>
        /// <value>The initial directory.</value>
        public string InitialDirectory
        {
            get { return this.fileDialog.InitialDirectory; }

            set { this.fileDialog.InitialDirectory = value; }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Gets the name of the file.
        /// </summary>
        public void GetFileName()
        {
            IntPtr ptr = GetForegroundWindow();
            var window = new WindowWrapper(ptr);
            if (this.fileDialog.ShowDialog(window) != DialogResult.OK)
            {
                this.fileDialog.FileName = string.Empty;
            }
        }

        /// <summary>
        /// Gets the foreground window.
        /// </summary>
        /// <returns>The foreground window</returns>
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        #endregion

        #region Sub Classes
        /// <summary>
        /// Instance of the WindowWrapper class.
        /// </summary>
        /// <author>Frank Alunni</author>
        internal class WindowWrapper : IWin32Window
        {
            #region Fields
            /// <summary>
            /// Window handle field
            /// </summary>
            private readonly IntPtr handle;
            #endregion

            #region Constructors
            /// <summary>
            /// Initializes a new instance of the <see cref="WindowWrapper"/> class.
            /// </summary>
            /// <param name="windowHandle">The window handle.</param>
            public WindowWrapper(IntPtr windowHandle)
            {
                this.handle = windowHandle;
            }
            #endregion

            #region Methods
            /// <summary>
            /// Gets the handle to the window represented by the implementer.
            /// </summary>
            /// <value></value>
            /// <returns>A handle to the window represented by the implementer.</returns>
            public virtual IntPtr Handle
            {
                get { return this.handle; }
            }
            #endregion
        }
        #endregion
    }
}
