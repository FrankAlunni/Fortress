// --------------------------------------------------------------------------------------------------------------------
// <copyright file="StatusbarChangedEventArgs.cs" company="B1C Canada Inc.">
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
    public class StatusbarChangedEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StatusbarChangedEventArgs"/> class.
        /// </summary>
        /// <param name="progressMessage">The progress message.</param>
        /// <param name="isErrorMessage">if set to <c>true</c> [is error message].</param>
        public StatusbarChangedEventArgs(string progressMessage, bool isErrorMessage)
        {
            this.ProgressMessage = progressMessage;
            this.IsErrorMessage = isErrorMessage;
        }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is error message.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is error message; otherwise, <c>false</c>.
        /// </value>
        public bool IsErrorMessage
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
