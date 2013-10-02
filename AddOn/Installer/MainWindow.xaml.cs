#region Usings
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;

using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using System.Threading;
#endregion

namespace B1C.Installer
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
    {
        #region Fields
        private InstallManager manager;

        /// <summary>
        /// An Empty delegate
        /// </summary>
        private static Action EmptyDelegate = delegate() { };
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class.
        /// </summary>
        public MainWindow()
		{
			this.InitializeComponent();

			// Insert code required on object creation below this point.
        }
        #endregion Constructors

        #region Events
        /// <summary>
        /// Handles the Click event of the createArdButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.RoutedEventArgs"/> instance containing the event data.</param>
        private void createArdButton_Click(object sender, RoutedEventArgs e)
        {
            manager.GenerateARD();
            this.Close();
        }

        /// <summary>
        /// Handles the Click event of the closeButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.RoutedEventArgs"/> instance containing the event data.</param>
        private void closeButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Handles the ContentRendered event of the Window control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void Window_ContentRendered(object sender, EventArgs e)
        {
            this.InitializeInstaller();
        }

        /// <summary>
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="StatusbarChangedEventArgs"/> instance containing the event data.</param>
        private void StatusChanged(object sender, StatusChangedEventArgs e)
        {
            this.waitLabel.Content = e.ProgressMessage;
            this.Refresh(this.waitLabel);
            System.Threading.Thread.Sleep(500);
        }

        #endregion Events

        /// <summary>
        /// Initializes the installer.
        /// </summary>
        private void InitializeInstaller()
        {

            manager = new InstallManager();
            manager.StatusChanged += this.StatusChanged;

            try
            {
                string commandLine = null; // The whole command line
                string[] commandLineElements = new string[3];

                // The command line parameters, seperated by '|' will be broken to this array
                int numberOfElements = 0; // The number of parameters in the command line (should be 2)

                numberOfElements = Environment.GetCommandLineArgs().Length;

                if (numberOfElements == 2)
                {
                    commandLine = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));

                    // Check if it is an unistall
                    if (commandLine.ToUpper() == "/U")
                    {
                        // Set Labels
                        this.closeButton.Visibility = System.Windows.Visibility.Hidden;
                        this.waitLabel.Content = string.Format("Uninstalling {0} version {1}", manager.InstallerInfo.ApplicationName, manager.InstallerInfo.ApplicationVersion);
                        this.productLabel.Content = string.Format("Thank you for using {0}. www.b1Computing.com", manager.InstallerInfo.ApplicationName);
                        this.Refresh(this.waitLabel);
                        this.Refresh(this.productLabel);
                        manager.UnInstall();
                    }

                    commandLineElements = commandLine.Split(char.Parse("|"));

                    // Get Install destination Folder
                    manager.DestinationPath = Convert.ToString(commandLineElements.GetValue(0));

                    // Get the "AddOnInstallAPI.dll" path
                    manager.DllPath = Convert.ToString(commandLineElements.GetValue(1));
                    manager.DllPath = manager.DllPath.Remove((manager.DllPath.Length - 19), 19); // Only the path is needed

                    // Hide the Close button
                    this.closeButton.Visibility = System.Windows.Visibility.Hidden;
                    this.createArdButton.Visibility = System.Windows.Visibility.Hidden; 
                }
                else
                {
                    this.closeButton.Visibility = System.Windows.Visibility.Visible;
                    this.createArdButton.Visibility = System.Windows.Visibility.Visible;
                }
            }
            catch (Exception ex)
            {
                this.closeButton.Visibility = System.Windows.Visibility.Visible;
                this.createArdButton.Visibility = System.Windows.Visibility.Hidden;
                manager.ShowError(ex);
            }

            this.productLabel.Content = string.Format("{0} version {1}", manager.InstallerInfo.ApplicationName, manager.InstallerInfo.ApplicationVersion);

            if (!string.IsNullOrEmpty(manager.DestinationPath))
            {
                manager.StartInstall();
            }
            else
            {
                this.waitLabel.Content = string.Format("Installer for {0} version {1}", manager.InstallerInfo.ApplicationName, manager.InstallerInfo.ApplicationVersion);
            }
            
        }

        /// <summary>
        /// Refreshes the specified UI element.
        /// </summary>
        /// <param name="uiElement">The UI element.</param>
        public void Refresh(UIElement uiElement)
        {
            uiElement.Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate);
        }
	}
}