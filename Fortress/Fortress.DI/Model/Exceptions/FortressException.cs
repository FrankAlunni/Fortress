//-----------------------------------------------------------------------
// <copyright file="FortressException.cs" company="Fortress Technology Inc.">
//     Copyright (c) Fortress Technology Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
//-----------------------------------------------------------------------

namespace Fortress.DI.Model.Exceptions
{
    #region Using Directive(s)

    using System;

    #endregion Using Directive(s)

    /// <summary>
    /// Base Exception for all Fortress namespace exceptions
    /// </summary>
    public class FortressException : ApplicationException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FortressException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        public FortressException(string message)
            : base(message)
        { }

        /// <summary>
        /// Initializes a new instance of the <see cref="FortressException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="innerException">The inner exception.</param>
        public FortressException(string message, Exception innerException)
            : base(message, innerException)
        { }
    }
}
