// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SingleApplication.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the SingleApplication type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------
namespace B1C.Utility.Windows
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Windows.Forms;

    /// <summary>
    /// Single Instance Application
    /// </summary>
    public class SingleApplication
    {
        /// <summary>
        /// The Restore flag
        /// </summary>
        private const int SwRestore = 9;

        /// <summary>
        /// Mutex class
        /// </summary>
        private static Mutex mutex;

        /// <summary>
        /// Execute a form base application if another instance already running on
        /// the system activate previous one
        /// </summary>
        /// <param name="frmMain">The main form</param>
        /// <returns>True if no previous instance is running</returns>
        public static bool Run(Form frmMain)
        {
            if (IsAlreadyRunning())
            {
                // Set focus on previously running app
                SwitchToCurrentInstance();
                return false;
            }

            Application.Run(frmMain);
            return true;
        }

        /// <summary>
        /// for console base application
        /// </summary>
        /// <returns>True if no previous instance is running</returns>
        public static bool Run()
        {
            return !IsAlreadyRunning();
        }

        /// <summary>
        /// Runs the specified assembly location.
        /// </summary>
        /// <param name="assemblyLocation">The assembly location.</param>
        /// <returns>True if no previous instance is running</returns>
        public static bool Run(string assemblyLocation)
        {
            if (!IsAlreadyRunning(assemblyLocation))
            {
                Application.Run();
                return true;
            }

            return false;
        }

        /// <summary>
        /// Shows the window.
        /// </summary>
        /// <param name="handleWindow">The handle window.</param>
        /// <param name="cmdShow">The coomand show.</param>
        /// <returns>The handle to the windows application</returns>
        [DllImport("user32.dll")]
        private static extern int ShowWindow(IntPtr handleWindow, int cmdShow);

        /// <summary>
        /// Sets the foreground window.
        /// </summary>
        /// <param name="handleWindow">The handle window.</param>
        /// <returns>The handle to the windows application</returns>
        [DllImport("user32.dll")]
        private static extern int SetForegroundWindow(IntPtr handleWindow);

        /// <summary>
        /// Determines whether the specified handle window is iconic.
        /// </summary>
        /// <param name="handleWindow">The handle window.</param>
        /// <returns>The handle to the icon of the application</returns>
        [DllImport("user32.dll")]
        private static extern int IsIconic(IntPtr handleWindow);

        /// <summary>
        /// Gets the current instance window handle.
        /// </summary>
        /// <returns>The handle to the windows application</returns>
        private static IntPtr GetCurrentInstanceWindowHandle()
        {
            IntPtr handle = IntPtr.Zero;
            Process process = Process.GetCurrentProcess();
            Process[] processes = Process.GetProcessesByName(process.ProcessName);
            foreach (Process currentProcess in processes)
            {
                // Get the first instance that is not this instance, has the
                // same process name and was started from the same file name
                // and location. Also check that the process has a valid
                // window handle in this session to filter out other user's
                // processes.
                if ((currentProcess.MainModule != null) && (process.MainModule != null))
                {
                    if (currentProcess.Id != process.Id && currentProcess.MainModule.FileName == process.MainModule.FileName &&
                        currentProcess.MainWindowHandle != IntPtr.Zero)
                    {
                        handle = currentProcess.MainWindowHandle;
                        break;
                    }
                }
            }

            return handle;
        }

        /// <summary>
        /// Switches to current instance.
        /// </summary>
        private static void SwitchToCurrentInstance()
        {
            IntPtr handle = GetCurrentInstanceWindowHandle();
            if (handle != IntPtr.Zero)
            {
                // Restore window if minimised. Do not restore if already in
                // normal or maximised window state, since we don't want to
                // change the current state of the window.
                if (IsIconic(handle) != 0)
                {
                    ShowWindow(handle, SwRestore);
                }

                // Set foreground window.
                SetForegroundWindow(handle);
            }
        }

        /// <summary>
        /// check if given exe alread running or not
        /// </summary>
        /// <returns>returns true if already running</returns>
        private static bool IsAlreadyRunning()
        {
            return IsAlreadyRunning(Assembly.GetExecutingAssembly().Location);
        }

        /// <summary>
        /// Determines whether [is already running] [the specified assembly location].
        /// </summary>
        /// <param name="assemblyLocation">The assembly location.</param>
        /// <returns>
        /// <c>true</c> if [is already running] [the specified assembly location]; otherwise, <c>false</c>.
        /// </returns>
        private static bool IsAlreadyRunning(string assemblyLocation)
        {
            bool createdNew = false;

            if (assemblyLocation != null)
            {
                FileSystemInfo fileInfo = new FileInfo(assemblyLocation);
                string exeName = fileInfo.Name;

                mutex = new Mutex(true, "Global\\" + exeName, out createdNew);
                if (createdNew)
                {
                    mutex.ReleaseMutex();
                }
            }

            return !createdNew;
        }
    }
}
