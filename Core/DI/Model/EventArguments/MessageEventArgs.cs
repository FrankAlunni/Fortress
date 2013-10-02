// --------------------------------------------------------------------------------------------------------------------
// <copyright file="MessageEventArgs.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   The Message Event
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.Model.EventArguments
{
    #region Using Directive(s)

    using System;

    using B1C.SAP.DI.Constants;

    #endregion Using Directive(s)

    public class MessageEventArgs : EventArgs
    {
        /// <summary>
        /// Gets the Level of the message being raised
        /// </summary>
        public MessageLevel MessageLevel
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the actual message being raised
        /// </summary>
        public string Message
        {
            get;
            private set;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MessageEventArgs"/> class.
        /// </summary>
        /// <param name="messageLevel">The message level.</param>
        /// <param name="message">The message.</param>
        public MessageEventArgs(MessageLevel messageLevel, string message)
        {
            this.MessageLevel = messageLevel;
            this.Message = message;
        }
    }
}
