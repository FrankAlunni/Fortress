// -----------------------------------------------------------------------
// <copyright file="PoolException.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <summary>
//    Product        : Messaging Service
//    Author Email   : bryan.atkinson@B1C.com
//    Created Date   : 08/26/2009
//    Prerequisites  : - SAP Business One v.2007a
//                     - SQL Server 2005
//                     - Micorsoft .NET Framework 3.5
// </summary>
//-----------------------------------------------------------------------

namespace B1C.Utility.Pools
{
    #region Using Directives

    using System;

    #endregion Using Directives

    /// <summary>
    /// This is the base exception for any Pool related exception
    /// </summary>
    public class PoolException : ApplicationException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PoolException"/> class.
        /// </summary>
        /// <param name="message">
        /// The message indicating the cause of the exception
        /// </param>
        public PoolException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PoolException"/> class.
        /// </summary>
        /// <param name="message">
        /// The message indicating the cause of the exception
        /// </param>
        /// <param name="innerException">The parent to this exception </param>
        public PoolException(string message, Exception innerException)
            : base(message, innerException)
        {            
        }
    }
}