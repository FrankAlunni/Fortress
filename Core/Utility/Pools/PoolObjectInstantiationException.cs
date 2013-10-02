﻿// -----------------------------------------------------------------------
// <copyright file="PoolObjectInstantiationException.cs" company="B1C Canada Inc.">
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
    /// <summary>
    /// This Exception should be thrown when an IObjectPool's cannot instantiate
    /// an object for inclusion in the pool.
    /// </summary>
    public class PoolObjectInstantiationException : PoolException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PoolObjectInstantiationException"/> class.
        /// </summary>
        /// <param name="message">
        /// The message indicating the cause of the exception
        /// </param>
        public PoolObjectInstantiationException(string message)
            : base(message)
        {
        }
    }
}