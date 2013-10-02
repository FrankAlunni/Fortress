//-----------------------------------------------------------------------
// <copyright file="DaysOfWeek.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------
namespace B1C.Utility.Enums
{
    #region Using Directive(s)

    using System;

    #endregion Using Directive(s)

    [Flags]
    public enum DaysOfWeek
    {
        None = 0,
        Sunday = 1,
        Monday = 2,
        Tuesday = 4,
        Wednesday = 8,
        Thursday = 16,
        Friday = 32,
        Saturday = 64
    }
}
