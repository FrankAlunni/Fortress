// -----------------------------------------------------------------------
// <copyright file="PoolFullException.cs" company="B1C Canada Inc.">
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
    /// This Exception should be thrown when an IObjectPool's acquire() method is 
    /// called, and the pool is full, but all objects are in use.
    /// When thrown, this exception indicates that either the pool's configured Maximum 
    /// Size is too small for the demands of application, or the calling classes are not 
    /// properly returning the acquired objects to the pool.
    /// </summary>
    public class PoolFullException : PoolException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PoolFullException"/> class.
        /// </summary>
        /// <param name="message">
        /// The message indicating the cause of the exception
        /// </param>
        public PoolFullException(string message) : base(message)
        {            
        }
    }
}