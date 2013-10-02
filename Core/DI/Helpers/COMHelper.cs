// --------------------------------------------------------------------------------------------------------------------
// <copyright file="COMHelper.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the COMHelper type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;
using SAPbobsCOM;

namespace B1C.SAP.DI.Helpers
{
    /// <summary>
    /// The COM Helper Class
    /// </summary>
    public static class COMHelper
    {
        #region Release Implementations
        /// <summary>
        /// Releases the specified COM object.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="comObject">The COM object.</param>
        public static void Release<T>(ref T comObject) where T : class
        {
            if (comObject == null) return;

            Marshal.ReleaseComObject(comObject);
            comObject = default(T);
            GC.Collect();
        }

        /// <summary>
        /// Releases the specified COM object without setting to null.
        /// </summary>
        /// <param name="comObject">The COM object.</param>
        public static void ReleaseOnly<T>(T comObject) where T : class
        {
            if (comObject == null) return;

            Marshal.ReleaseComObject(comObject);
            GC.Collect();
        }
        #endregion Release Implementations

        #region Other Generic Methods

        /// <summary>
        /// Sets the users defined field value.
        /// </summary>
        /// <param name="userFields">The user fields.</param>
        /// <param name="index">The index.</param>
        /// <param name="fieldValue">The field value.</param>
        public static void UserDefinedFieldValue(UserFields userFields, object index, object fieldValue)
        {
            Fields fields = userFields.Fields;
            try
            {
                Field field = fields.Item(index);
                try
                {
                    field.Value = fieldValue;
                }
                finally
                {
                    Release(ref field);
                }
            }
            finally
            {
                Release(ref fields);
                Release(ref userFields);
            }
        }

        /// <summary>
        /// Reutns the fields value.
        /// </summary>
        /// <param name="userFields">The user fields.</param>
        /// <param name="index">Index of the field.</param>
        /// <returns>The field value</returns>
        public static object UserDefinedFieldValue(UserFields userFields, object index)
        {
            Fields fields = userFields.Fields;

            try
            {
                Field field = fields.Item(index);

                try
                {
                    return field.Value;
                }
                finally
                {
                    Release(ref field);
                }
            }
            finally
            {
                Release(ref fields);
                Release(ref userFields);
            }
        }

        #endregion Other Generic Methods
    }
}