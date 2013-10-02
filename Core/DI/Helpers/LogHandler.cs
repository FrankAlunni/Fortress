//-----------------------------------------------------------------------
// <copyright file="LogHandler.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.DI.Helpers
{
    using System;
    using System.Collections;
    using System.IO;
    using B1C.SAP.DI.Components;

    /// <summary>
    /// Instance of the LogHandler class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public static class LogHandler
    {
        #region Fields
        /// <summary>
        /// True if you want to log information messages
        /// </summary>
        private static bool logInfo;

        /// <summary>
        /// True is it is initialized
        /// </summary>
        private static bool isInitialized;

        private static string LogPath
        {
            get
            {
                string logPath = string.Empty;

                ConfigurationHelper.Load("LogPath", out logPath);

                if (string.IsNullOrEmpty(logPath))
                {
                    logPath = System.Windows.Forms.Application.StartupPath;
                }

                return logPath;
            }
        }
        #endregion Fields

        #region Public Methods

        /// <summary>
        /// Gets the last logged error.
        /// </summary>
        /// <returns>The last logged error</returns>
        public static string GetLastLoggedError()
        {
            try
            {
                string filePath = string.Format(
                    @"{0}\Errors\{1}_{2}{3}{4}.log", 
                    LogPath,
                    System.Windows.Forms.Application.ProductName, 
                    DateTime.Today.Year,
                    DateTime.Today.Month.ToString("00"),
                    DateTime.Today.Day.ToString("00"));

                const string New_line = "===================== START ERROR SUMMARY =========================";
                const string End_line = "====================== END ERROR SUMMARY ==========================";
                char[] newLineChar = "~".ToCharArray();
                string errorFile;
                using (StreamReader rdr = File.OpenText(filePath))
                {
                    errorFile = rdr.ReadToEnd();
                    errorFile = errorFile.Replace(New_line, "~");
                }

                string[] errorLogged = errorFile.Split(newLineChar);
                if (errorLogged.Length == 1)
                {
                    return errorLogged[0];
                }
                
                if (errorLogged.Length > 1)
                {
                    return errorLogged[errorLogged.Length - 2].Replace(End_line, string.Empty);
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="ex">The exception.</param>
        public static void LogError(Exception ex)
        {
            try
            {
                CheckFolder();
                CheckFile();
                FileStream fileStream = CreateStream();
                StreamWriter streamWriter = CreateWriter(fileStream);
                WriteException(streamWriter, ex);
                CloseStreamAndWriter(fileStream, streamWriter);
            }
            catch 
            { 
            }
        }

        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="ex">The exception.</param>
        /// <param name="message">The message.</param>
        public static void LogError(Exception ex, string message)
        {
            try
            {
                CheckFolder();
                CheckFile();
                FileStream fileStream = CreateStream();
                StreamWriter streamWriter = CreateWriter(fileStream);
                WriteExceptionWithMessage(streamWriter, ex, message);
                CloseStreamAndWriter(fileStream, streamWriter);
            }
            catch 
            { 
            }
        }

        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="ex">The exception.</param>
        /// <param name="messages">The messages.</param>
        public static void LogError(Exception ex, ArrayList messages)
        {
            try
            {
                CheckFolder();
                CheckFile();
                FileStream fileStream = CreateStream();
                StreamWriter streamWriter = CreateWriter(fileStream);
                WriteExceptionWithMessages(streamWriter, ex, messages);
                CloseStreamAndWriter(fileStream, streamWriter);
            }
            catch 
            { 
            }
        }

        /// <summary>
        /// Logs the information.
        /// </summary>
        /// <param name="message">The message.</param>
        public static void LogInformation(string message)
        {
            try
            {
                if (!logInfo)
                {
                    return;
                }

                CheckFolder();
                CheckFile();
                FileStream fileStream = CreateStream();
                StreamWriter streamWriter = CreateWriter(fileStream);
                WriteMessage(streamWriter, message);
                CloseStreamAndWriter(fileStream, streamWriter);
            }
            catch 
            {
            }
        }

        /// <summary>
        /// Logs the information.
        /// </summary>
        /// <param name="messages">The messages.</param>
        public static void LogInformation(ArrayList messages)
        {
            try
            {
                if (!logInfo)
                {
                    return;
                }

                CheckFolder();
                CheckFile();
                FileStream fileStream = CreateStream();
                StreamWriter streamWriter = CreateWriter(fileStream);
                WriteMessages(streamWriter, messages);
                CloseStreamAndWriter(fileStream, streamWriter);
            }
            catch 
            { 
            }
        }
        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Initializes the logger.
        /// </summary>
        private static void InitializeLogger()
        {
            if (isInitialized)
            {
                return;
            }

            ConfigurationHelper.Load("LogInformation", out logInfo);
            isInitialized = true;
        }

        /// <summary>
        /// Checks the file.
        /// </summary>
        private static void CheckFile()
        {
            try
            {
                string filePath = string.Format(
                    @"{0}\Errors\{1}_{2}{3}{4}.log",
                    LogPath,
                    System.Windows.Forms.Application.ProductName, 
                    DateTime.Today.Year,
                    DateTime.Today.Month.ToString("00"),
                    DateTime.Today.Day.ToString("00"));

                var fileStream = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                var streamWriter = new StreamWriter(fileStream);
                CloseStreamAndWriter(fileStream, streamWriter);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Checks the folder.
        /// </summary>
        private static void CheckFolder()
        {
            InitializeLogger();

            if (!Directory.Exists(string.Format(@"{0}\Errors\", LogPath)))
            {
                Directory.CreateDirectory(string.Format(@"{0}\Errors\", LogPath));
            }
        }

        /// <summary>
        /// Closes the stream and writer.
        /// </summary>
        /// <param name="fileStream">The fileStream.</param>
        /// <param name="streamWriter">The streamWriter.</param>
        private static void CloseStreamAndWriter(Stream fileStream, TextWriter streamWriter)
        {
            streamWriter.Close();
            fileStream.Close();
            streamWriter.Dispose();
            fileStream.Dispose();
        }

        /// <summary>
        /// Creates the stream.
        /// </summary>
        /// <returns>The file stream</returns>
        private static FileStream CreateStream()
        {
            FileStream fileStream;
            try
            {
                string filePath = string.Format(
                    @"{0}\Errors\{1}_{2}{3}{4}.log",
                    LogPath,
                    System.Windows.Forms.Application.ProductName,
                    DateTime.Today.Year,
                    DateTime.Today.Month.ToString("00"),
                    DateTime.Today.Day.ToString("00"));

                fileStream = new FileStream(filePath, FileMode.Append, FileAccess.Write);
            }
            catch
            {
                fileStream = null;
            }

            return fileStream;
        }

        /// <summary>
        /// Creates the writer.
        /// </summary>
        /// <param name="fileStream">The fileStream.</param>
        /// <returns>The stream writer.</returns>
        private static StreamWriter CreateWriter(Stream fileStream)
        {
            var streamWriter = new StreamWriter(fileStream);
            return streamWriter;
        }

        /// <summary>
        /// Writes the exception.
        /// </summary>
        /// <param name="streamWriter">The streamWriter.</param>
        /// <param name="ex">The exception.</param>
        private static void WriteException(TextWriter streamWriter, Exception ex)
        {
            streamWriter.WriteLine("===================== START ERROR SUMMARY =========================");
            streamWriter.WriteLine(string.Format("Source        ==> {0}", ex.Source));
            streamWriter.WriteLine(string.Format("HelpLink      ==> {0}", ex.HelpLink));
            streamWriter.WriteLine(string.Format("Message       ==> {0}", ex.Message));
            string separator = Environment.NewLine;
            char[] sepChars = separator.ToCharArray();
            string[] stack = ex.StackTrace.Split(sepChars);
            for (int indexStack = 0; indexStack < stack.Length; indexStack++)
            {
                if (stack[indexStack].Trim() != string.Empty)
                {
                    if (indexStack == 0)
                    {
                        WriteStackTraceLine(
                            streamWriter,
                            "Stack Trace   ==> {0}",
                            "                  in {0}",
                            stack[indexStack].Trim());
                    }
                    else
                    {
                        WriteStackTraceLine(
                            streamWriter,
                            "                  {0}",
                            "                  in {0}",
                            stack[indexStack].Trim());
                    }
                }
            }

            streamWriter.WriteLine(string.Format("Target Site   ==> {0}", ex.TargetSite));
            streamWriter.WriteLine(string.Format("Date/Time     ==> {0}", DateTime.Now));
            streamWriter.WriteLine("====================== END ERROR SUMMARY ==========================");
        }

        /// <summary>
        /// Writes the exception.
        /// </summary>
        /// <param name="streamWriter">The streamWriter.</param>
        /// <param name="ex">The exception.</param>
        /// <param name="message">The message.</param>
        private static void WriteExceptionWithMessage(TextWriter streamWriter, Exception ex, string message)
        {
            streamWriter.WriteLine("===================== START ERROR SUMMARY =========================");
            streamWriter.WriteLine(string.Format("Source        ==> {0}", ex.Source));
            streamWriter.WriteLine(string.Format("HelpLink      ==> {0}", ex.HelpLink));
            streamWriter.WriteLine(string.Format("Message       ==> {0}", ex.Message));
            string separator = Environment.NewLine;
            char[] sepChars = separator.ToCharArray();
            string[] stack = ex.StackTrace.Split(sepChars);
            for (int indexStack = 0; indexStack < stack.Length; indexStack++)
            {
                if (stack[indexStack].Trim() != string.Empty)
                {
                    if (indexStack == 0)
                    {
                        WriteStackTraceLine(
                            streamWriter,
                            "Stack Trace   ==> {0}",
                            "                  in {0}",
                            stack[indexStack].Trim());
                    }
                    else
                    {
                        WriteStackTraceLine(
                            streamWriter,
                            "                  {0}",
                            "                  in {0}",
                            stack[indexStack].Trim());
                    }
                }
            }

            WriteMessage(streamWriter, message);
            streamWriter.WriteLine(string.Format("Target Site   ==> {0}", ex.TargetSite));
            streamWriter.WriteLine(string.Format("Date/Time     ==> {0}", DateTime.Now));
            streamWriter.WriteLine("====================== END ERROR SUMMARY ==========================");
        }

        /// <summary>
        /// Writes the exception.
        /// </summary>
        /// <param name="streamWriter">The streamWriter.</param>
        /// <param name="ex">The exception.</param>
        /// <param name="messages">The messages.</param>
        private static void WriteExceptionWithMessages(TextWriter streamWriter, Exception ex, IList messages)
        {
            streamWriter.WriteLine("===================== START ERROR SUMMARY =========================");
            streamWriter.WriteLine(string.Format("Source        ==> {0}", ex.Source));
            streamWriter.WriteLine(string.Format("HelpLink      ==> {0}", ex.HelpLink));
            streamWriter.WriteLine(string.Format("Message       ==> {0}", ex.Message));
            string separator = Environment.NewLine;
            char[] sepChars = separator.ToCharArray();
            string[] stack = ex.StackTrace.Split(sepChars);
            for (int indexStack = 0; indexStack < stack.Length; indexStack++)
            {
                if (stack[indexStack].Trim() != string.Empty)
                {
                    if (indexStack == 0)
                    {
                        WriteStackTraceLine(
                            streamWriter,
                            "Stack Trace   ==> {0}",
                            "                  in {0}",
                            stack[indexStack].Trim());
                    }
                    else
                    {
                        WriteStackTraceLine(
                            streamWriter,
                            "                  {0}",
                            "                  in {0}",
                            stack[indexStack].Trim());
                    }
                }
            }

            WriteMessages(streamWriter, messages);
            streamWriter.WriteLine(string.Format("Target Site   ==> {0}", ex.TargetSite));
            streamWriter.WriteLine(string.Format("Date/Time     ==> {0}", DateTime.Now));
            streamWriter.WriteLine("====================== END ERROR SUMMARY ==========================");
        }

        /// <summary>
        /// Writes the stack trace line.
        /// </summary>
        /// <param name="streamWriter">The streamWriter.</param>
        /// <param name="formatString">The format string.</param>
        /// <param name="formatStringSubLine">The format string sub line.</param>
        /// <param name="message">The message.</param>
        private static void WriteStackTraceLine(TextWriter streamWriter, string formatString, string formatStringSubLine, string message)
        {
            string[] lines = message.Split(new[] { " in " }, 99, StringSplitOptions.RemoveEmptyEntries);
            for (int indexLine = 0; indexLine < lines.Length; indexLine++)
            {
                if (lines[indexLine].Trim() != string.Empty)
                {
                    switch (indexLine)
                    {
                        case 0:
                            streamWriter.WriteLine(string.Format(formatString, lines[indexLine].Trim()));
                            break;
                        default:
                            streamWriter.WriteLine(string.Format(formatStringSubLine, lines[indexLine].Trim()));
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Writes the message.
        /// </summary>
        /// <param name="streamWriter">The streamWriter.</param>
        /// <param name="message">The message.</param>
        private static void WriteMessage(TextWriter streamWriter, string message)
        {
            streamWriter.WriteLine(string.Format("Message       ==> {0}", message));
        }

        /// <summary>
        /// Writes the messages.
        /// </summary>
        /// <param name="streamWriter">The streamWriter.</param>
        /// <param name="messages">The messages.</param>
        private static void WriteMessages(TextWriter streamWriter, IList messages)
        {
            if (messages != null)
            {
                for (int indexList = 0; indexList < messages.Count; indexList++)
                {
                    switch (indexList)
                    {
                        case 0:
                            streamWriter.WriteLine(string.Format("Message       ==> {0}", messages[indexList]));
                            break;
                        default:
                            streamWriter.WriteLine(string.Format("                  {0}", messages[indexList]));
                            break;
                    }
                }
            }
        }

        #endregion Private Methods
    }
}
