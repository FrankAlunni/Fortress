//-----------------------------------------------------------------------
// <copyright file="LockManager.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------

namespace B1C.Utility.Threading
{
    #region Using Directive(s)

    using System;
    using System.Configuration;
    using Enums;
    using Logging;

    #endregion Using Directive(s)

    public static class LockManager
    {
        /// <summary>
        /// A flag indicating whether or not the lock level has been read from the configuration
        /// </summary>
        private static bool _lockLevelRead;

        /// <summary>
        /// The actual Lock Level
        /// </summary>
        private static LockLevel _lockLevel;

        /// <summary>
        /// The locking object
        /// </summary>
        private static readonly object LockObject = new object();

        /// <summary>
        /// Internal Locking object
        /// </summary>
        private static readonly object InnerLock = new object();

        /// <summary>
        /// Gets the current lock level.
        /// </summary>
        /// <value>The current lock level.</value>
        private static LockLevel CurrentLockLevel
        {
            get
            {
                if (!_lockLevelRead)
                {
                    lock (InnerLock)
                    {
                        if (!_lockLevelRead)
                        {
                            string level = ConfigurationManager.AppSettings["LockLevel"];

                            if (!string.IsNullOrEmpty(level))
                            {
                                if (Enum.IsDefined(typeof(LockLevel), level))
                                {
                                    _lockLevel = (LockLevel)Enum.Parse(typeof(LockLevel), level);
                                }
                                else
                                {
                                    ThreadedAppLog.WriteLine(
                                        "Unable to read LockLevel of \"{0}\". Defaulting to no locking.", level);
                                    _lockLevel = LockLevel.NoLock;
                                }
                            }
                            else
                            {
                                ThreadedAppLog.WriteLine(
                                    "No lock level found in configuration. Defaulting to no locking");
                                _lockLevel = LockLevel.NoLock;
                            }
                        }

                        _lockLevelRead = true;
                    }
                }

                return _lockLevel;
            }
        }

        /// <summary>
        /// Gets the lock (if level required = system lock level).
        /// </summary>
        /// <param name="levelRequired">The level required.</param>
        /// <returns>A lock object</returns>
        public static object GetLock(LockLevel levelRequired)
        {
            if (levelRequired == CurrentLockLevel)
            {
                return LockObject;
            }

            return new object();
        }

    }

}
