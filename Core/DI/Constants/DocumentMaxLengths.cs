using System;
using System.Collections.Generic;

using System.Text;

namespace B1C.SAP.DI.Constants
{
    public static class DocumentMaxLengths
    {
        /// <summary>
        /// Document Comments field maximum length
        /// </summary>
        public const int Comments = 254;

        public static string TrimLength(string value, int fieldLength)
        {
            string newValue = value;

            if (newValue.Length > fieldLength)
            {
                newValue = newValue.Substring(0, fieldLength - 3) + "...";
            }

            return newValue;
        }
    
    }
}
