// -----------------------------------------------------------------------
// <copyright file="IObjectPool.cs" company="B1C Canada Inc.">
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
    /// This interface defines all the public methods/properties of a Pool
    /// of objects
    /// </summary>
    /// <typeparam name="T">
    /// The type of object that the pool distributes
    /// </typeparam>
    public interface IObjectPool<T>
    {
        /// <summary>
        /// Gets the name of the pool
        /// </summary>
        string PoolName { get;  }

        /// <summary>
        /// Gets the maximum number of instantiated objects in the pool
        /// </summary>
        int MaxObjects { get; }

        /// <summary>
        /// Acquires an object from the pool
        /// </summary>
        /// <returns>
        /// An object pulled from the pool
        /// </returns>
        T Acquire();

        /// <summary>
        /// Releases an in-use object, back into the pool for use by others
        /// </summary>
        /// <param name="obj">The object being returned</param>
        void Release(T obj);
    }
}