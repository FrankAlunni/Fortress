//-----------------------------------------------------------------------
// <copyright file="MonthsOfYear.cs" company="B1C Canada Inc.">
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
    public enum MonthsOfYear
    {
        None = 0,
        January = 1,
        February = 2,
        March = 4,
        April = 8,
        May = 16,
        June = 32,
        July = 64,
        August = 128,
        September = 256,
        October = 512,
        November = 1024,
        December = 2048
    }
}
