// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NamedAppLog.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Class to log application messages.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.Utility.Logging
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using B1C.Utility.Helpers;
    using System.Configuration;

    /// <summary>
    /// Class to log application messages.
    /// </summary>
    public class NamedAppLog
    {
        #region Fields

        /// <summary>
        /// The lock object
        /// </summary>
        private static readonly object LockObject = new object();

        /// <summary>
        /// Message Line
        /// </summary>
        private const string MsgLine = "*******************************************************************************";

        /// <summary>
        /// Message Header
        /// </summary>
        private const string MsgAppheader = "Started at {0}";

        /// <summary>
        /// Message Foother
        /// </summary>
        private const string MsgAppfooter = "Ended at {0}";

        /// <summary>
        /// Message time stamp line
        /// </summary>
        private const string MsgTimestampedline = "{0:MM}/{0:dd}/{0:yyyy} {0:HH}:{0:mm}:{0:ss}.{0:fff}: {1}";

        /// <summary>
        /// The working threaded tasks
        /// </summary>
        private static IDictionary<string, IList<string>> workingThreadTasks = new Dictionary<string, IList<string>>();

        /// <summary>
        /// Working threaded output
        /// </summary>
        private static IDictionary<string, StringBuilder> workingThreadOutput = new Dictionary<string, StringBuilder>();

        /// <summary>
        /// Working threaded delayed writes
        /// </summary>
        private static IDictionary<string, bool> workingThreadDelayedWrite = new Dictionary<string, bool>();
        #endregion

        #region Constructors

        /// <summary>
        /// Prevents a default instance of the <see cref="NamedAppLog"/> class from being created. 
        /// </summary>
        private NamedAppLog()
        {
            ConsoleOutput = true;
            workingThreadOutput = new Dictionary<string, StringBuilder>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets a value indicating whether [console output].
        /// </summary>
        /// <value><c>true</c> if [console output]; otherwise, <c>false</c>.</value>
        public static bool ConsoleOutput { get; set; }

        /// <summary>
        /// Gets or sets the logs dir.
        /// </summary>
        /// <value>The logs dir.</value>
        public static string LogsDir { get; set; }       

        #endregion

        #region Static Methods

        /// <summary>
        /// Opens the log.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        public static void OpenLog(string threadName)
        {
            if (string.IsNullOrEmpty(LogsDir))
            {
                try
                {
                    LogsDir = ConfigurationManager.AppSettings["LogsDir"];
                }
                catch { }
            }

            workingThreadOutput[threadName] = new StringBuilder(string.Empty);
            workingThreadTasks[threadName] = new List<string>();
            workingThreadDelayedWrite[threadName] = true;
        }

        public static void OpenLog(string threadName, bool delayedWrite)
        {
            if (string.IsNullOrEmpty(LogsDir))
            {
                try
                {
                    LogsDir = ConfigurationManager.AppSettings["LogsDir"];
                }
                catch { }
            }

            workingThreadOutput[threadName] = new StringBuilder(string.Empty);
            workingThreadTasks[threadName] = new List<string>();
            workingThreadDelayedWrite[threadName] = delayedWrite;
        }

        /// <summary>
        /// Closes the log.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <param name="writeLog">if set to <c>true</c> [write log].</param>
        public static void CloseLog(string threadName, bool writeLog)
        {
            FlushLog(threadName, writeLog);
        }

        /// <summary>
        /// Gets the log.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <returns>The log file</returns>
        public static string GetLog(string threadName)
        {
            if (!workingThreadOutput.ContainsKey(threadName))
            {
                return string.Empty;
            }

            var strOutput = workingThreadOutput[threadName];

            return strOutput.ToString();
        }

        /// <summary>
        /// Writes the line.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <param name="format">The format.</param>
        /// <param name="arg">The arguments.</param>
        public static void WriteLine(string threadName, string format, params object[] arg)
        {
            if (!workingThreadOutput.ContainsKey(threadName))
            {
                return;
            }
            
            StringBuilder strOutput = workingThreadOutput[threadName];

            string s;
            if (arg != null && arg.Length != 0)
            {
                s = string.Format(format, arg) + "\r\n";
            }
            else
            {
                s = format + "\r\n";
            }

            strOutput.Append(s);
            if (ConsoleOutput)
            {
                Console.Write(s);
            }

            if (workingThreadDelayedWrite.ContainsKey(threadName) && !workingThreadDelayedWrite[threadName])
            {
                FlushLog(threadName, true);
            }
        }

        /// <summary>
        /// Writes the specified thread name.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <param name="format">The format.</param>
        /// <param name="arg">The arguments.</param>
        public static void Write(string threadName, string format, params object[] arg)
        {
            if (!workingThreadOutput.ContainsKey(threadName))
            {
                return;
            }
            
            StringBuilder strOutput = workingThreadOutput[threadName];

            string s;
            if (arg != null && arg.Length != 0)
            {
                s = string.Format(format, arg);
            }
            else
            {
                s = format;
            }

            strOutput.Append(s);
            if (ConsoleOutput)
            {
                Console.Write(s);
            }

            if (workingThreadDelayedWrite.ContainsKey(threadName) && !workingThreadDelayedWrite[threadName])
            {
                FlushLog(threadName, true);
            }
        }

        /// <summary>
        /// Writes the blank line.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        public static void WriteBlankLine(string threadName)
        {
            WriteLine(threadName, string.Empty);
        }

        /// <summary>
        /// Writes the time stamped line.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <param name="format">The format.</param>
        /// <param name="arg">The arguments.</param>
        public static void WriteTimeStampedLine(string threadName, string format, params object[] arg)
        {
            if (!workingThreadOutput.ContainsKey(threadName))
            {
                return;
            }
            
            StringBuilder strOutput = workingThreadOutput[threadName];

            string s;
            if (arg != null && arg.Length != 0)
            {
                s = string.Format(MsgTimestampedline, DateTime.Now, string.Format(format, arg)) + "\r\n";
            }
            else
            {
                s = string.Format(MsgTimestampedline, DateTime.Now, format) + "\r\n";
            }

            strOutput.Append(s);
            if (ConsoleOutput)
            {
                Console.Write(s);
            }

            if (workingThreadDelayedWrite.ContainsKey(threadName) && !workingThreadDelayedWrite[threadName])
            {
                FlushLog(threadName, true);
            }
        }

        /// <summary>
        /// Writes the error line.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <param name="error">The error.</param>
        public static void WriteErrorLine(string threadName, string error)
        {
            WriteBlankLine(threadName);
            WriteLine(threadName, "Error: " + error);
        }

        /// <summary>
        /// Writes the error line.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <param name="error">The error.</param>
        /// <param name="task">The task running.</param>
        public static void WriteErrorLine(string threadName, string error, string task)
        {
            WriteBlankLine(threadName);
            WriteLine(threadName, "Error executing task " + task + " : " + error);
        }

        /// <summary>
        /// Writes the error line.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <param name="e">The exception.</param>
        public static void WriteErrorLine(string threadName, Exception e)
        {
            WriteBlankLine(threadName);
            WriteLine(threadName, "Exception Source: " + e.Source);
            WriteLine(threadName, "Exception Message: " + e.Message);
            WriteLine(threadName, "Exception Stack Trace: " + e.StackTrace);
        }

        /// <summary>
        /// Writes the error line.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <param name="e">The exception.</param>
        /// <param name="task">The task running.</param>
        public static void WriteErrorLine(string threadName, Exception e, string task)
        {
            WriteBlankLine(threadName);
            WriteLine(threadName, "Error executing task: " + task);
            WriteLine(threadName, "Exception Source: " + e.Source);
            WriteLine(threadName, "Exception Message: " + e.Message);
            WriteLine(threadName, "Exception Stack Trace: " + e.StackTrace);
        }

        /// <summary>
        /// Writes the app header.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        public static void WriteAppHeader(string threadName)
        {
            WriteLine(threadName, MsgLine);
            WriteLine(threadName, MsgAppheader + " (" + threadName + ")", DateTime.Now);
            WriteLine(threadName, MsgLine);
            WriteBlankLine(threadName);
        }

        /// <summary>
        /// Writes the app header.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <param name="applicationName">Name of the application.</param>
        public static void WriteAppHeader(string threadName, string applicationName)
        {
            WriteLine(threadName, MsgLine);
            WriteLine(threadName, applicationName + " " + MsgAppheader + " (" + threadName + ")", DateTime.Now);
            WriteLine(threadName, MsgLine);
            WriteBlankLine(threadName);
        }

        /// <summary>
        /// Writes the app footer.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        public static void WriteAppFooter(string threadName)
        {
            WriteBlankLine(threadName);
            WriteLine(threadName, MsgLine);
            WriteLine(threadName, MsgAppfooter, DateTime.Now);
            WriteLine(threadName, MsgLine);
            WriteBlankLine(threadName);
        }

        /// <summary>
        /// Writes the app footer.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <param name="applicationName">Name of the application.</param>
        public static void WriteAppFooter(string threadName, string applicationName)
        {
            WriteBlankLine(threadName);
            WriteLine(threadName, MsgLine);
            WriteLine(threadName, applicationName + " " + MsgAppfooter, DateTime.Now);
            WriteLine(threadName, MsgLine);
            WriteBlankLine(threadName);
        }

        /// <summary>
        /// Gets the log between.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <returns>The last section of log between the given characters</returns>
        public static string GetLogBetween(DateTime date, string start, string end)
        {
            string logMessage = string.Empty;

            string currentFile = GetLogNameFromDate(date);
            string[] searchResults = FileHelper.Search(currentFile, start, end);
            if (searchResults.Length > 0)
            {
                logMessage = searchResults[searchResults.Length - 1];
            }

            return logMessage;
        }

        /// <summary>
        /// Gets the log between.
        /// </summary>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <returns>The last section of log between the given characters</returns>
        public static string GetLogBetween(string start, string end)
        {
            return GetLogBetween(DateTime.Now, start, end);
        }

        /// <summary>
        /// Gets the log between.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <param name="toReplace">To replace.</param>
        /// <param name="toSubstitute">To substitute.</param>
        /// <returns>The last section of log between the given characters</returns>
        public static string GetLogBetween(DateTime date, string start, string end, string toReplace, string toSubstitute)
        {
            string logMessage = GetLogBetween(date, start, end);
            return logMessage.Replace(toReplace, toSubstitute);
        }

        /// <summary>
        /// Gets the log between.
        /// </summary>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <param name="toReplace">To replace.</param>
        /// <param name="toSubstitute">To substitute.</param>
        /// <returns>The last section of log between the given characters</returns>
        public static string GetLogBetween(string start, string end, string toReplace, string toSubstitute)
        {
            return GetLogBetween(DateTime.Now, start, end, toReplace, toSubstitute);
        }

        /// <summary>
        /// Sets the task.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <param name="task">The task.</param>
        public static void SetTask(string threadName, string task)
        {
            if (workingThreadTasks.ContainsKey(threadName))
            {
                IList<string> taskList = workingThreadTasks[threadName];
                if (taskList == null)
                {
                    taskList = new List<string>();
                    workingThreadTasks[threadName] = taskList;
                }

                workingThreadTasks[threadName].Add(task);
            }

            WriteLine(threadName, task);
        }

        /// <summary>
        /// Gets the current task.
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <returns>The current task</returns>
        public static string GetCurrentTask(string threadName)
        {
            string task = string.Empty;
            if (workingThreadTasks.ContainsKey(threadName))
            {
                IList<string> taskList = workingThreadTasks[threadName];
                if (taskList != null)
                {
                    task = taskList[taskList.Count - 1];
                }
            }

            return task;
        }

        /// <summary>
        /// Gets the task list for the given thread
        /// </summary>
        /// <param name="threadName">Name of the thread.</param>
        /// <returns>The task list for the given thread</returns>
        public static IList<string> GetTaskList(string threadName)
        {
            IList<string> taskList = null;
            if (workingThreadTasks.ContainsKey(threadName))
            {
                taskList = workingThreadTasks[threadName];
            }

            if (taskList == null)
            {
                taskList = new List<string>();
            }

            return taskList;
        }

        #endregion

        private static void FlushLog(string threadName, bool writeLog)
        {
            if (!workingThreadOutput.ContainsKey(threadName) || !writeLog)
            {
                return;
            }

            StringBuilder strOutput = workingThreadOutput[threadName];
            TextWriter textWriter = null;

            lock (LockObject)
            {
                try
                {
                    string fileName = string.Format("{0:yyyy}-{0:MM}-{0:dd}", DateTime.Now);

                    try
                    {
                        textWriter = new StreamWriter(Path.Combine((LogsDir != string.Empty ? LogsDir : string.Empty), fileName) + ".log", true);
                    }
                    catch
                    {
                        textWriter = new StreamWriter(fileName + ".log", true);
                    }

                    textWriter.Write(strOutput.ToString());
                }
                finally
                {
                    if (textWriter != null)
                    {
                        textWriter.Close();
                    }
                }
            }

            workingThreadOutput[threadName] = new StringBuilder();
        }

        private static string GetLogNameFromDate(DateTime date)
        {
            string fileName = string.Format("{0:yyyy}-{0:MM}-{0:dd}", date);

            return (Path.Combine((LogsDir != string.Empty ? LogsDir : string.Empty), fileName) + ".log");
        }

        private static string CurrentLogName
        {
            get
            {
                string fileName = string.Format("{0:yyyy}-{0:MM}-{0:dd}", DateTime.Now);

                return (Path.Combine((LogsDir != string.Empty ? LogsDir : string.Empty), fileName) + ".log");
            }
        }
    }
}