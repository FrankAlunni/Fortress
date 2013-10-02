// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ProgressChangedEventArgs.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   The Progress Event
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.Model.EventArguments
{
    using System;

    /// <summary>
    /// The Progress Event
    /// </summary>
    public class ProgressChangedEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ProgressChangedEventArgs"/> class.
        /// </summary>
        /// <param name="progressValue">The progress value.</param>
        /// <param name="progressMessage">The progress message.</param>
        public ProgressChangedEventArgs(int progressValue, string progressMessage)
        {
            this.ProgressValue = progressValue;
            this.ProgressMessage = progressMessage;
        }

        /// <summary>
        /// Gets or sets the progress value.
        /// </summary>
        /// <value>The progress value.</value>
        public int ProgressValue
        {
            get;
            set;
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
