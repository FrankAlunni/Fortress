// -----------------------------------------------------------------------
// <copyright file="BaseObjectPool.cs" company="B1C Canada Inc.">
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

    using System.Collections.Generic;

    #endregion Using Directives

    /// <summary>
    /// This is the base implementation of a pool of objects
    /// </summary>
    /// <typeparam name="T">
    /// The type of object the pool distributes
    /// </typeparam>
    public abstract class BaseObjectPool<T> : IObjectPool<T>
    {
        #region Private Members

        /// <summary>
        /// The pool of available objects
        /// </summary>
        private static readonly List<T> AvailableObjects = new List<T>();

        /// <summary>
        /// The objects currently in use from the pool
        /// </summary>
        private static readonly List<T> UnavailableObjects = new List<T>();

        /// <summary>
        /// The object used for class level synchronziation
        /// </summary>
        private static readonly object LockObject = new object();

        #endregion Private Members

        #region Public Method(s)

        #region IObjectPool<T> Members

        /// <summary>
        /// Gets the maximum number of instantiated objects in the pool
        /// </summary>
        public int MaxObjects { get; private set; }

        /// <summary>
        /// Gets the name of the pool
        /// </summary>
        public abstract string PoolName { get; }

        /// <summary>
        /// Gets the current size of the pool (both available and in-use objects)
        /// </summary>
        public int PoolSize
        {
            get
            {
                return AvailableObjects.Count + UnavailableObjects.Count; 
            }
        }

        /// <summary>
        /// Acquires an object from the pool
        /// </summary>
        /// <returns>
        /// An object pulled from the pool
        /// </returns>
        public virtual T Acquire()
        {
            T availableObject;

            // This method modifies the pool of available objects. Need to synchronize access to this
            // pool to prevent multiple acquisitions of the same object
            lock (LockObject)
            {
                // If there are available objects, return the first
                // Otherwise, create a new object if there is capacity
                if (AvailableObjects.Count > 0)
                {
                    availableObject = AvailableObjects[0];

                    // The object is in use, so move it to the list of unavailable objects
                    AvailableObjects.RemoveAt(0);
                    UnavailableObjects.Add(availableObject);
                }
                else
                {
                    // If there is no capacity for more objects, throw a NoObjectAvailable exception
                    // Otherwise create a new object and add it to the list of unavailable objects
                    // A 0 value for MaxObjects indicates that there is no limit
                    if (this.PoolSize < this.MaxObjects || this.MaxObjects == 0)
                    {
                        availableObject = this.CreateObject();

                        // TO be done: Could the object returned be null?
                        UnavailableObjects.Add(availableObject);
                    }
                    else
                    {
                        throw new PoolFullException(string.Format("The pool [{0}] is full. Max Size: {1}. Available Objects {2}. Objects in use: {3}.", this.PoolName, this.MaxObjects, AvailableObjects.Count, UnavailableObjects.Count));
                    }
                }
            }

            return availableObject;
        }

        /// <summary>
        /// Releases an in-use object, back into the pool for use by others
        /// </summary>
        /// <param name="obj">The object being returned</param>
        public virtual void Release(T obj)
        {
            // This method modifies the pool of available objects. Need to synchronize access to this
            // pool to prevent multiple acquisitions of the same object
            lock (LockObject)
            {
                // Check if the object being released was in fact checked out
                if (UnavailableObjects.Contains(obj))
                {
                    // Move the object from the unavailable list to the available list
                    UnavailableObjects.Remove(obj);
                    AvailableObjects.Add(obj);
                }
            }
        }

        /// <summary>
        /// Disposes of an object
        /// </summary>
        /// <param name="obj">The object to dispose of</param>
        public virtual void DisposeObject(T obj)
        {
            // Removes the object from the queue
            if (AvailableObjects.Contains(obj))
            {
                AvailableObjects.Remove(obj);
            }

            if (UnavailableObjects.Contains(obj))
            {
                UnavailableObjects.Remove(obj);
            }
        }

        #endregion

        #endregion Public Method(s)

        #region Protected Method(s)

        /// <summary>
        /// Creates an object of type T
        /// </summary>
        /// <returns>
        /// A newly instantiated object of type T
        /// </returns>
        protected abstract T CreateObject();

        #endregion Protected Method(s)
    }
}