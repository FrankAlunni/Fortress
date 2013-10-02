using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace B1C.SAP.DI.Model.Exceptions
{
    public class SapImportException : ApplicationException
    {
        /// <summary>
        /// The inner exception
        /// </summary>
        private readonly Exception _innerException;

        /// <summary>
        /// Initializes a new instance of the <see cref="SapImportException"/> class.
        /// </summary>
        /// <param name="ex">The ex.</param>
        public SapImportException(Exception ex)
        {
            this._innerException = ex;
        }

        /// <summary>
        /// Gets a message that describes the current exception.
        /// </summary>
        /// <value></value>
        /// <returns>
        /// The error message that explains the reason for the exception, or an empty string("").
        /// </returns>
        public override string Message
        {
            get
            {
                return "SAP Import Exception: " + _innerException.Message;
            }
        }

        /// <summary>
        /// Gets a string representation of the frames on the call stack at the time the current exception was thrown.
        /// </summary>
        /// <value></value>
        /// <returns>
        /// A string that describes the contents of the call stack, with the most recent method call appearing first.
        /// </returns>
        /// <PermissionSet>
        /// 	<IPermission class="System.Security.Permissions.FileIOPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" PathDiscovery="*AllFiles*"/>
        /// </PermissionSet>
        public override string StackTrace
        {
            get
            {
                return _innerException.StackTrace;
            }
        }
    }
}
