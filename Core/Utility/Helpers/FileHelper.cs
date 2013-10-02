// --------------------------------------------------------------------------------------------------------------------
// <copyright file="FileHelper.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Help manipulate and retreive file information
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.Utility.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// Help manipulate and retreive file information
    /// </summary>
    public static class FileHelper
    {
        /// <summary>
        /// Returns the file name
        /// </summary>
        /// <param name="fullPath">Full path for a file, including folder names.</param>
        /// <returns>The extension for the file.</returns>
        /// <example>
        /// The following code
        /// <code>
        ///     string fullName = @"C:\Projects\Demo\LFile.cs";
        ///     string result = LFile.Extension(fullPath);
        /// </code>
        /// will return <c>LFile</c> as result.
        /// </example>
        public static string FileName(string fullPath)
        {
            if (fullPath != null)
            {
                int pos = fullPath.LastIndexOf(@"\");

                if (pos > 0)
                {
                    return fullPath.Substring(pos + 1);
                }
            }

            return fullPath;
        }

        /// <summary>
        /// Returns the extension for the file name
        /// </summary>
        /// <param name="fullPath">Full path for a file, including folder names.</param>
        /// <returns>The extension for the file.</returns>
        /// <example>
        /// The following code
        /// <code>
        ///     string fullPath = @"C:\Projects\Demo\LFile.cs";
        ///     string result = LFile.Extension(fullPath);
        /// </code>
        /// will return <c>cs</c> as result.
        /// </example>
        public static string FileExtension(string fullPath)
        {
            if (!string.IsNullOrEmpty(fullPath))
            {
                string fileName = FileName(fullPath);

                int pos = fileName.LastIndexOf(".");

                if (pos > 0)
                {
                    return fileName.Substring(pos + 1);
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Searches the specified full path.
        /// </summary>
        /// <param name="fullPath">The full path.</param>
        /// <param name="beginString">The begin string.</param>
        /// <param name="endString">The end string.</param>
        /// <returns>A list of the strings found between the beginning and the end string.</returns>
        public static string[] Search(string fullPath, string beginString, string endString)
        {
            var results = new List<string>();
            StreamReader reader = File.OpenText(fullPath);
            string strSource;
            bool foundOpenTag = false;
            string result = string.Empty;
            while ((strSource = reader.ReadLine()) != null)
            {
                int indexOfBegin;

                if (foundOpenTag)
                {
                    indexOfBegin = strSource.IndexOf(endString);
                    if (indexOfBegin > -1)
                    {
                        result += strSource;
                        results.Add(result);
                        result = string.Empty;
                        foundOpenTag = false;
                    }
                }
                else
                {
                    indexOfBegin = strSource.IndexOf(beginString);

                    if (indexOfBegin > -1)
                    {
                        foundOpenTag = true;
                    }
                }

                if (!foundOpenTag)
                {
                    continue;
                }

                result += strSource;
                result += Environment.NewLine;
            }

            reader.Close();

            return results.ToArray();
        }
    }
}
