// --------------------------------------------------------------------------------------------------------------------
// <copyright file="MessageLevel.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   The Message Event
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.Constants
{
    #region Using Directive(s)

    using System;

    #endregion Using Directive(s)

    /// <summary>
    /// Defines the level of a logging message
    /// </summary>
    public enum MessageLevel
    {
        /// <summary>
        /// Debug level messages
        /// </summary>
        DEBUG,
        /// <summary>
        /// Information only messages
        /// </summary>
        INFO,
        /// <summary>
        /// A warning of a potential problem
        /// </summary>
        WARN,
        /// <summary>
        /// An error has occurred
        /// </summary>
        ERROR,
        /// <summary>
        /// A fatal error has occurred, and the system must be shut down
        /// </summary>
        FATAL
    }
}
