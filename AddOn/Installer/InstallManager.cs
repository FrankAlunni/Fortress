using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using System.Threading;
using System.Diagnostics;

namespace B1C.Installer
{
    public class InstallManager
    {

        /// <summary>
        /// The File Watcher
        /// </summary>
        private System.IO.FileSystemWatcher FileWatcher;

        /// <summary>
        /// The resources namespace
        /// </summary>
        public string ResourceNamespace;

        /// <summary>
        /// The resources namespace
        /// </summary>
        public string SharedResourceNamespace;

        /// <summary>
        /// The path of "AddOnInstallAPI.dll"
        /// </summary>
        public string DllPath;

        /// <summary>
        /// Installation target path
        /// </summary>
        public string DestinationPath;

        /// <summary>
        /// True if the file was created
        /// </summary>
        private bool isFileCreated;

        /// <summary>
        /// 
        /// </summary>
        public readonly InstallerInfo InstallerInfo;

        /// <summary>
        /// Registration exe path
        /// </summary>
        private string registrationEXEPath;

        #region Delegates
        /// <summary>
        /// The handler for the statusbar changed
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The StatusChangedEventArgs arguments.</param>
        public delegate void StatusChangedEventHandler(object sender, StatusChangedEventArgs e);
        #endregion Delegates

        #region Events
        /// <summary>
        /// Occurs when [statusbar change].
        /// </summary>
        public event StatusChangedEventHandler StatusChanged;
        #endregion Events

        public InstallManager()
        {

            this.DllPath = @"C:\Program Files\SAP\SAP Business One";
            this.ResourceNamespace = this.GetType().Namespace + ".Resources.";
            this.SharedResourceNamespace = this.GetType().Namespace + ".SharedResources."; ;
            this.InstallerInfo = new InstallerInfo();

            // Initialize the File watcher
            this.FileWatcher = new System.IO.FileSystemWatcher();
            this.FileWatcher.Renamed += new System.IO.RenamedEventHandler(this.FileWatcher_Renamed);
        }


        /// <summary>
        /// Raises the StatusbarChanged event.
        /// </summary>
        /// <param name="e">The <see cref="Maritz.SAP.DI.Components.StatusChanged"/> instance containing the event data.</param>
        public virtual void OnStatusChanged(StatusChangedEventArgs e)
        {
            this.StatusChanged(this, e);
        }

        /// <summary>
        /// Handles the Renamed event of the FileWatcher control.
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The event arguments.</param>
        private void FileWatcher_Renamed(object sender, RenamedEventArgs e)
        {
            this.isFileCreated = true;
            this.FileWatcher.EnableRaisingEvents = false;
        }


        /// <summary>
        /// Starts the install.
        /// </summary>
        public void StartInstall()
        {
            try
            {
                Environment.CurrentDirectory = this.DllPath; // For Dll function calls will work

                if (!Directory.Exists(this.DestinationPath))
                {
                    Directory.CreateDirectory(this.DestinationPath); // Create installation folder
                }

                this.FileWatcher.Path = this.DestinationPath;
                this.FileWatcher.EnableRaisingEvents = true;

                // Extract add-on to installation folder
                foreach (string file in this.InstallerInfo.AddonFileNames)
                {
                    if (this.StatusChanged != null)
                    {
                        this.OnStatusChanged(new StatusChangedEventArgs(string.Format("Extracting {0} ...", file.Replace(this.ResourceNamespace, string.Empty))));
                    }
                    this.ExtractFile(this.DestinationPath, file, this.ResourceNamespace);
                }

                // Extract add-on to installation folder
                string sapFolder = this.DestinationPath.Substring(0, this.DestinationPath.IndexOf(@"\AddOns"));
                //sapFolder = this.DestinationPath;
                foreach (string file in this.InstallerInfo.AddonSharedFileNames)
                {
                    if (this.StatusChanged != null)
                    {
                        this.OnStatusChanged(new StatusChangedEventArgs(string.Format("Extracting {0} ...", file.Replace(this.SharedResourceNamespace, string.Empty))));
                    }
                    this.ExtractFile(sapFolder, file, this.SharedResourceNamespace);
                }


                if (this.StatusChanged != null)
                {
                    this.OnStatusChanged(new StatusChangedEventArgs("Installation completed successfully. Closing Installer."));
                }

                // Write installation Folder to registry
                this.WriteToRegistry(RegistryHive.LocalMachine, "SOFTWARE", this.InstallerInfo.ApplicationName, this.DestinationPath);

                EndInstall();
            }
            catch (Exception ex)
            {
                this.ShowError(ex);
            }

            this.ExitInstaller();
        }


        /// <summary>
        /// Extracts the file.
        /// </summary>
        /// <param name="path">The current path.</param>
        /// <param name="filename">The filename.</param>
        private void ExtractFile(string path, string filename, string nameSpace)
        {
            try
            {
                FileStream addonExeFile = null;
                Assembly thisExe = null;
                Stream file = null;

                // Enable Raise of event for completion of extract
                thisExe = Assembly.GetExecutingAssembly();

                object sourcePath = path + @"\" + filename + ".tmp";
                object targetPath = path + @"\" + filename.Replace(nameSpace, string.Empty);

                file = thisExe.GetManifestResourceStream(filename);

                if (file == null)
                {
                    throw new Exception(string.Format("Could not find file {0} in the installer resources.", filename));
                }

                // Create a tmp file first, after file is extracted change to exe
                if (File.Exists(Convert.ToString(sourcePath)))
                {
                    File.Delete(Convert.ToString(sourcePath));
                }

                addonExeFile = File.Create(Convert.ToString(sourcePath));

                if (file != null)
                {
                    byte[] buffer = new byte[file.Length];

                    file.Read(buffer, 0, Convert.ToInt32(file.Length));
                    addonExeFile.Write(buffer, 0, Convert.ToInt32(file.Length));
                }

                addonExeFile.Close();

                if (File.Exists(Convert.ToString(targetPath)))
                {
                    File.Delete(Convert.ToString(targetPath));
                }

                // Change file extension to exe
                File.Move(Convert.ToString(sourcePath), Convert.ToString(targetPath));

                // Don't continue running until the file is copied...
                while (this.isFileCreated == false)
                {
                    this.DoEvents();
                }

                this.isFileCreated = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Occurred");
                this.isFileCreated = false;
            }
        }

        /// <summary>
        /// Uns the install.
        /// </summary>
        public void UnInstall()
        {
            string path = null;
            path = this.ReadPath(); // Reads the addon path from the registry

            if (string.IsNullOrEmpty(path))
            {
                path = this.FindFolder(this.DllPath, this.InstallerInfo.ApplicationName);
            }
            this.DeleteFolder(path);

            // Terminate the application
            EndUninstall(string.Empty, true);

            GC.Collect();
            Environment.Exit(0);
        }

        /// <summary>
        /// Read the addon path from the registry.
        /// </summary>
        /// <returns>The addon path</returns>
        private string ReadPath()
        {
            string readPathReturn = null;
            string answer = null;
            string error = string.Empty;

            answer = this.RegValue(RegistryHive.LocalMachine, "SOFTWARE", this.InstallerInfo.ApplicationName, ref error);
            readPathReturn = answer;
            //if (answer == string.Empty)
            //{
            //    MessageBox.Show("This error occurred: " + error);
            //}

            return readPathReturn;
        }

        /// <summary>
        /// This Function reads values to the registry
        /// </summary>
        /// <param name="hive">The registry hive.</param>
        /// <param name="key">The registry key.</param>
        /// <param name="valueName">The value name.</param>
        /// <param name="errInfo">The error info.</param>
        /// <returns>The registry value</returns>
        public string RegValue(RegistryHive hive, string key, string valueName, [Optional] ref string errInfo)
        {
            RegistryKey objParent = null;
            RegistryKey objSubkey = null;
            string answer = null;
            switch (hive)
            {
                case RegistryHive.ClassesRoot:
                    objParent = Registry.ClassesRoot;
                    break;
                case RegistryHive.CurrentConfig:
                    objParent = Registry.CurrentConfig;
                    break;
                case RegistryHive.CurrentUser:
                    objParent = Registry.CurrentUser;
                    break;
                case RegistryHive.DynData:
                    objParent = Registry.DynData;
                    break;
                case RegistryHive.LocalMachine:
                    objParent = Registry.LocalMachine;
                    break;
                case RegistryHive.PerformanceData:
                    objParent = Registry.PerformanceData;
                    break;
                case RegistryHive.Users:
                    objParent = Registry.Users;

                    break;
            }

            try
            {
                if (objParent != null)
                {
                    objSubkey = objParent.OpenSubKey(key);
                }

                // if can't be found, object is not initialized
                if (!(objSubkey == null))
                {
                    answer = Convert.ToString(objSubkey.GetValue(valueName));
                }
            }
            catch (Exception ex)
            {
                errInfo = ex.Message;
            }
            finally
            {
                // if no error but value is empty, populate errinfo
                if (errInfo == string.Empty & answer == string.Empty)
                {
                    errInfo = "No value found for requested registry key";
                }
            }

            return answer;
        }

        /// <summary>
        /// Deletes the folder.
        /// </summary>
        /// <param name="path">The current path.</param>
        private void DeleteFolder(string path)
        {
            if (Directory.Exists(path))
            {
                string[] files = Directory.GetFiles(path);
                string[] dirs = Directory.GetDirectories(path);
                foreach (string file in files)
                {
                    File.SetAttributes(file, FileAttributes.Normal);
                    File.Delete(file);
                }

                foreach (string dir in dirs)
                {
                    this.DeleteFolder(dir);
                }

                Directory.Delete(path, false);
            }
        }

        /// <summary>
        /// Signals SBO that the uninstall is complete.
        /// </summary>
        /// <param name="str">The path of the unistall.</param>
        /// <param name="b">if set to <c>true</c> is is completed successfully.</param>
        /// <returns>0 for failure 1 for success </returns>
        [DllImport("AddOnInstallAPI.dll")]
        private static extern int EndUninstall(string str, bool b);

        /// <summary>
        /// Signals SBO that the installation is complete.
        /// </summary>
        /// <returns>0 for failure 1 for success </returns>
        [DllImport("AddOnInstallAPI.dll")]
        private static extern int EndInstall();

        /// <summary>
        /// Exits the installer.
        /// </summary>
        private void ExitInstaller()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // terminating the Add On
            Environment.Exit(0);
        }

        /// <summary>
        /// Shows the error.
        /// </summary>
        /// <param name="ex">The exception.</param>
        public void ShowError(Exception ex)
        {
            MessageBox.Show(ex.Message + Environment.NewLine + "Source:" + ex.StackTrace);
        }

        /// <summary>
        /// Writes to registry.
        /// </summary>
        /// <param name="parentKeyHive">The parent key hive.</param>
        /// <param name="subKeyName">The sub key name.</param>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value.</param>
        /// <returns>True if it writes to the registry</returns>
        private bool WriteToRegistry(RegistryHive parentKeyHive, string subKeyName, string valueName, object value)
        {
            RegistryKey objSubKey = null;
            RegistryKey objParentKey = null;

            try
            {
                switch (parentKeyHive)
                {
                    case RegistryHive.ClassesRoot:
                        objParentKey = Registry.ClassesRoot;
                        break;
                    case RegistryHive.CurrentConfig:
                        objParentKey = Registry.CurrentConfig;
                        break;
                    case RegistryHive.CurrentUser:
                        objParentKey = Registry.CurrentUser;
                        break;
                    case RegistryHive.DynData:
                        objParentKey = Registry.DynData;
                        break;
                    case RegistryHive.LocalMachine:
                        objParentKey = Registry.LocalMachine;
                        break;
                    case RegistryHive.PerformanceData:
                        objParentKey = Registry.PerformanceData;
                        break;
                    case RegistryHive.Users:
                        objParentKey = Registry.Users;
                        break;
                }

                // Open 
                if (objParentKey != null)
                {
                    objSubKey = objParentKey.OpenSubKey(subKeyName, true) ?? objParentKey.CreateSubKey(subKeyName);

                    // create if doesn't exist
                }

                if (objSubKey != null)
                {
                    objSubKey.SetValue(valueName, value);
                }
            }
            catch
            {
            }

            return true;
        }

        private void DoEvents()
        {
            DispatcherFrame f = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,
            (SendOrPostCallback)delegate(object arg)
            {
                DispatcherFrame fr = arg as DispatcherFrame;
                fr.Continue = false;
            }, f);
            Dispatcher.PushFrame(f);
        }

        /// <summary>
        /// Generates the ARD.
        /// </summary>
        public void GenerateARD()
        {
            string solutionFolder = System.IO.Directory.GetCurrentDirectory();
            string destinationFolder = solutionFolder + "\\Deployment";

            // string paramterXMLFileName = installerInfo.ApplicationName + ".xml";
            string installerEXEFileName = Assembly.GetExecutingAssembly().ManifestModule.Name;
            string paramterXMLFileName = installerEXEFileName.Replace(".exe", ".xml");
            string paramterXMLFile = destinationFolder + "\\" + paramterXMLFileName;

            // Create deployment folder
            if (Directory.Exists(destinationFolder))
            {
                Directory.Delete(destinationFolder, true);
            }

            Directory.CreateDirectory(destinationFolder);

            // MessageBox.Show("Folder Created");

            // Create a .xml file containing the information needed by the .ard file
            FileInfo fi = new FileInfo(paramterXMLFile);
            if (fi.Exists)
            {
                fi.Delete(); // if file exists delete it
            }

            // MessageBox.Show("Delete XML 1");
            StreamWriter sw = new StreamWriter(fi.Create());
            sw.Write(
                    "<AddOnInfo partnernmsp=" + PutQuote(this.InstallerInfo.PartnerNamespace) + " contdata="
                    + PutQuote(this.InstallerInfo.PartnerContactData) + " addonname="
                    + PutQuote(installerEXEFileName.Replace(".exe", string.Empty)) + " addongroup="
                    + PutQuote(this.InstallerInfo.AddOnGroup) + " esttime="
                    + PutQuote(this.InstallerInfo.EstInstTime) + " instparams="
                    + PutQuote(this.InstallerInfo.InstallParameters) + " uncmdarg="
                    + PutQuote(this.InstallerInfo.UninstallParameters) + " partnername="
                    + PutQuote(this.InstallerInfo.PartnerName) + " unesttime="
                    + PutQuote(this.InstallerInfo.EstUninstTime) + " />");
            sw.Close();
            sw.Dispose();

            // MessageBox.Show("Write XML");

            // Copy Installer into destination
            string exePath = (new Uri(Assembly.GetExecutingAssembly().GetName().CodeBase)).LocalPath;
            File.Copy(exePath, destinationFolder + "\\" + installerEXEFileName);

            // MessageBox.Show("Copy Installer From: " + Application.ExecutablePath);
            // MessageBox.Show("Copy Installer To: " + destinationFolder + "\\" + installerEXEFileName);

            // Copy main exe into destination
            this.FileWatcher.Path = destinationFolder;
            this.FileWatcher.EnableRaisingEvents = true;
            this.ExtractFile(destinationFolder, this.ResourceNamespace + this.InstallerInfo.AddOnExeFile, this.ResourceNamespace);
            this.FileWatcher.EnableRaisingEvents = false;

            if (!this.isFileCreated)
            {
                return;
            }

            // MessageBox.Show("Extract EXE: " + installerInfo.AddOnExeFile);
            this.registrationEXEPath = this.GetAddOnRegDataGenPath();

            // MessageBox.Show(registrationEXEPath);
            string args = paramterXMLFileName + " " + this.InstallerInfo.ApplicationVersion + " " + installerEXEFileName
                          + " " + installerEXEFileName + " " + this.InstallerInfo.AddOnExeFile + " ";

            // MessageBox.Show(args);

            // Launch .bat for the .ard registration file creation
            ProcessStartInfo processStartInfo = new ProcessStartInfo(this.registrationEXEPath)
            {
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                WorkingDirectory = destinationFolder,
                Arguments = args
            };
            Process process = Process.Start(processStartInfo);
            if (process != null)
            {
                process.WaitForExit(20000); // wait for 20 sec
                process.Close();
            }

            // Try to rename the addonname value
            string replaceWhat = string.Format("addonname=\"{0}\"", installerEXEFileName.Replace(".exe", string.Empty));
            string replaceWith = string.Format("addonname=\"{0}\"", this.InstallerInfo.ApplicationName);

            FileStream fs = new FileStream(paramterXMLFile.Replace(".xml", ".ard"), FileMode.Open, FileAccess.ReadWrite);
            StreamReader sr = new StreamReader(fs);
            string line = null;
            string newLine = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                newLine += line.Replace(replaceWhat, replaceWith);
            }

            fs.Close();
            sr.Close();

            if (!string.IsNullOrEmpty(newLine))
            {
                FileStream newArd = new FileStream(
                        paramterXMLFile.Replace(".xml", ".ard"), FileMode.Create, FileAccess.Write);
                StreamWriter newFile = new StreamWriter(newArd);
                newFile.Write(newLine);

                newFile.Flush();
                newFile.Close();

                newArd.Close();
            }

            // Delete XML File
            File.Delete(paramterXMLFile);

            // MessageBox.Show("Delete XML 1");

            // Delete addon exe file
            File.Delete(destinationFolder + "\\" + this.InstallerInfo.AddOnExeFile);

            // MessageBox.Show("Delete EXE");
            this.DoEvents();
        }


        /// <summary>
        /// Gets the add on reg data gen path.
        /// </summary>
        /// <returns>The path for the registration app exe</returns>
        private string GetAddOnRegDataGenPath()
        {
            string path = string.Empty;

            RegistryKey regParam = Registry.CurrentUser.OpenSubKey("SOFTWARE\\SAP\\SAP Business One Development Environment\\SAP Business One AddOnInstaller Wizards\\", true);
            if (regParam != null)
            {
                path = (string)regParam.GetValue("registrationEXEPath");
                regParam.Close();
            }

            if (path != null && File.Exists(path))
            {
                return path;
            }

            if (this.registrationEXEPath == string.Empty || !File.Exists(this.registrationEXEPath))
            {
                // Look for directory inside registry writen by B1DE installation
                RegistryKey businessOnePathKey =
                        Registry.CurrentUser.OpenSubKey("SOFTWARE\\SAP\\SAP Business One Development Environment\\", false);
                if (businessOnePathKey != null)
                {
                    this.registrationEXEPath = (string)businessOnePathKey.GetValue("b1Path") + " SDK\\Tools\\AddOnRegDataGen\\AddOnRegDataGen.exe";
                }
            }

            if (this.registrationEXEPath == string.Empty || !File.Exists(this.registrationEXEPath))
            {
                // Look for default directory inside Program Files
                string defRegDataGenPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\\SAP\\SAP Business One SDK\\Tools\\AddOnRegDataGen\\AddOnRegDataGen.exe";
                this.registrationEXEPath = File.Exists(defRegDataGenPath) ? defRegDataGenPath : string.Empty;
            }

            this.SetAddOnRegDataGenPath(this.registrationEXEPath);

            return this.registrationEXEPath;
        }

        /// <summary>
        /// Sets the add on reg data gen path.
        /// </summary>
        /// <param name="path">The add on path.</param>
        private void SetAddOnRegDataGenPath(string path)
        {
            this.registrationEXEPath = path;

            RegistryKey regParam = Registry.CurrentUser.OpenSubKey(
                    "SOFTWARE\\SAP\\SAP Business One Development Environment\\SAP Business One AddOnInstaller Wizards\\",
                    true);
            if (regParam != null)
            {
                regParam.SetValue("registrationEXEPath", this.registrationEXEPath);
                regParam.Close();
            }
        }

        /// <summary>
        /// Puts the quote.
        /// </summary>
        /// <param name="inString">The in string.</param>
        /// <returns>The input string with quotes</returns>
        private static string PutQuote(string inString)
        {
            return string.Format("\"{0}\"", inString);
        }

        private string FindFolder(string startFrom, string sDir)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(startFrom))
                {
                    if (d.EndsWith(sDir))
                    {
                        return d;
                    }
                }

                foreach (string d in Directory.GetDirectories(startFrom))
                {
                    string subfolder = this.FindFolder(d, sDir);

                    if (!string.IsNullOrEmpty(subfolder))
                    {
                        return subfolder;
                    }
                }

            }
            catch 
            {
            }

            return string.Empty;
        }


    }

    public class StatusChangedEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StatusChanged"/> class.
        /// </summary>
        /// <param name="progressMessage">The progress message.</param>
        /// <param name="isErrorMessage">if set to <c>true</c> [is error message].</param>
        public StatusChangedEventArgs(string progressMessage)
        {
            this.ProgressMessage = progressMessage;
        }

        /// <summary>
        /// Gets or sets the progress message.
        /// </summary>
        /// <value>The progress message.</value>
        public string ProgressMessage
        {
            get;
            set;
        }
    }

}
