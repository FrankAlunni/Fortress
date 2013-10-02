// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ThreadedAppLog.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the ThreadedAppLog type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.Utility.Logging
{
    using System;
    using System.Threading;
    using System.Collections.Generic;

    /// <summary>
    /// The Threaded App Log
    /// </summary>
    public class ThreadedAppLog
    {
        /// <summary>
        /// Gets or sets a value indicating whether [console output].
        /// </summary>
        /// <value><c>true</c> if [console output]; otherwise, <c>false</c>.</value>
        public static bool ConsoleOutput
        {
            get
            {
                return NamedAppLog.ConsoleOutput;
            }

            set
            {
                NamedAppLog.ConsoleOutput = value;
            }
        }

        /// <summary>
        /// Gets or sets the logs dir.
        /// </summary>
        /// <value>The logs dir.</value>
        public static string LogsDir
        {
            get
            {
                return NamedAppLog.LogsDir;
            }

            set
            {
                NamedAppLog.LogsDir = value;
            }
        }

        /// <summary>
        /// Opens the log.
        /// </summary>
        public static void OpenLog()
        {
            NamedAppLog.OpenLog(GetThreadName());
        }

        /// <summary>
        /// Opens the log.
        /// </summary>
        public static void OpenLog(bool delayedWrite)
        {
            NamedAppLog.OpenLog(GetThreadName(), delayedWrite);
        }

        /// <summary>
        /// Closes the log.
        /// </summary>
        /// <param name="writeLog">if set to <c>true</c> [write log].</param>
        public static void CloseLog(bool writeLog)
        {
            NamedAppLog.CloseLog(GetThreadName(), writeLog);
        }

        /// <summary>
        /// Gets the log.
        /// </summary>
        /// <returns>Gets the log</returns>
        public static string GetLog()
        {
            return NamedAppLog.GetLog(GetThreadName());
        }

        /// <summary>
        /// Writes the line.
        /// </summary>
        /// <param name="format">The format.</param>
        /// <param name="arg">The arguments.</param>
        public static void WriteLine(string format, params object[] arg)
        {
            NamedAppLog.WriteLine(GetThreadName(), format, arg);
        }

        /// <summary>
        /// Writes the specified format.
        /// </summary>
        /// <param name="format">The format.</param>
        /// <param name="arg">The arguments.</param>
        public static void Write(string format, params object[] arg)
        {
            NamedAppLog.Write(GetThreadName(), format, arg);
        }

        /// <summary>
        /// Writes the blank line.
        /// </summary>
        public static void WriteBlankLine()
        {
            NamedAppLog.WriteBlankLine(GetThreadName());
        }

        /// <summary>
        /// Writes the time stamped line.
        /// </summary>
        /// <param name="format">The format.</param>
        /// <param name="arg">The arguments.</param>
        public static void WriteTimeStampedLine(string format, params object[] arg)
        {
            NamedAppLog.WriteTimeStampedLine(GetThreadName(), format, arg);
        }

        /// <summary>
        /// Writes the error line.
        /// </summary>
        /// <param name="error">The error.</param>
        public static void WriteErrorLine(string error)
        {
            NamedAppLog.WriteErrorLine(GetThreadName(), error);
        }

        /// <summary>
        /// Writes the error line.
        /// </summary>
        /// <param name="error">The error.</param>
        /// <param name="task">The task running.</param>
        public static void WriteErrorLine(string error, string task)
        {
            NamedAppLog.WriteErrorLine(GetThreadName(), error, task);
        }

        /// <summary>
        /// Writes the error line.
        /// </summary>
        /// <param name="e">The exception.</param>
        public static void WriteErrorLine(Exception e)
        {
            NamedAppLog.WriteErrorLine(GetThreadName(), e);
        }

        /// <summary>
        /// Writes the error line.
        /// </summary>
        /// <param name="e">The exception.</param>
        /// <param name="task">The task running.</param>
        public static void WriteErrorLine(Exception e, string task)
        {
            NamedAppLog.WriteErrorLine(GetThreadName(), e, task);
        }

        /// <summary>
        /// Writes the app header.
        /// </summary>
        public static void WriteAppHeader()
        {
            NamedAppLog.WriteAppHeader(GetThreadName());
        }

        /// <summary>
        /// Writes the app header.
        /// </summary>
        /// <param name="applicationName">Name of the application.</param>
        public static void WriteAppHeader(string applicationName)
        {
            NamedAppLog.WriteAppHeader(GetThreadName(), applicationName);
        }

        /// <summary>
        /// Writes the app footer.
        /// </summary>
        public static void WriteAppFooter()
        {
            NamedAppLog.WriteAppFooter(GetThreadName());
        }

        /// <summary>
        /// Writes the app footer.
        /// </summary>
        /// <param name="applicationName">Name of the application.</param>
        public static void WriteAppFooter(string applicationName)
        {
            NamedAppLog.WriteAppFooter(GetThreadName(), applicationName);
        }

        /// <summary>
        /// Gets the name of the thread.
        /// </summary>
        /// <returns>The thread name</returns>
        protected static string GetThreadName()
        {
            string threadName = Thread.CurrentThread.Name ?? "DEFAULT";

            return threadName;
        }

        public static string GetLogBetween(DateTime date, string start, string end)
        {
            return NamedAppLog.GetLogBetween(date, start, end);
        }

        public static string GetLogBetween(string start, string end)
        {
            return NamedAppLog.GetLogBetween(start, end);
        }

        public static string GetLogBetween(string start, string end, string toReplace, string toSubstitute)
        {
            return NamedAppLog.GetLogBetween(start, end, toReplace, toSubstitute);
        }

        public static string GetLogBetween(DateTime date, string start, string end, string toReplace, string toSubstitute)
        {
            return NamedAppLog.GetLogBetween(date, start, end, toReplace, toSubstitute);
        }

        public static void SetTask(string task)
        {
            NamedAppLog.SetTask(GetThreadName(), task);
        }

        public static string GetCurrentTask()
        {
            return NamedAppLog.GetCurrentTask(GetThreadName());
        }

        public static IList<string> GetTaskList(string threadName)
        {
            return NamedAppLog.GetTaskList(GetThreadName());
        }

    }
}
