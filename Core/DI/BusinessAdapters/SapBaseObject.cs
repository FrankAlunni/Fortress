// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SapBaseObject.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   SAP base object
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters
{
    using System;

    /// <summary>
    /// SAP base object
    /// </summary>
    public abstract class SapBaseObject : IDisposable
    {
        /// <summary>
        /// Whether the object is disposed or not
        /// </summary>
        private bool disposed;

        /// <summary>
        /// Initializes a new instance of the <see cref="SapBaseObject"/> class.
        /// </summary>
        protected SapBaseObject()
        {
            this.disposed = false;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (!this.disposed)
            {
                try
                {
                    this.Release();
                }
                finally
                {
                    this.disposed = true;
                }
            }
        }

        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected abstract void Release();
    }
}
