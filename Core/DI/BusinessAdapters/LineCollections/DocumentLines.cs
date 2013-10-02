//-----------------------------------------------------------------------
// <copyright file="DocumentLines.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
using B1C.SAP.DI.Helpers;

namespace B1C.SAP.DI.BusinessAdapters
{
    using System;
    using System.Collections;
    using SAPbobsCOM;
    using Utility.Helpers;

    /// <summary>
    /// Instance of the Lines class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public class Lines : SapBaseObject, IEnumerable, IEnumerator
    {
        #region Fields
        /// <summary>
        /// The SAP document lines
        /// </summary>
        private Document_Lines lines;

        /// <summary>
        /// The current line
        /// </summary>
        private int currentLine;
        #endregion Fields

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="Lines"/> class.
        /// </summary>
        /// <param name="lines">The input lines.</param>
        public Lines(Document_Lines lines)
        {
            this.lines = lines;
        }

        /// <summary>
        /// Gets the count.
        /// </summary>
        /// <value>The line count.</value>
        public int Count
        {
            get { return this.lines.Count; }
        }
        #endregion Constructor

        #region Methods
        #region IEnumerator Members

        /// <summary>
        /// Gets the current element in the collection.
        /// </summary>
        /// <value></value>
        /// <returns>The current element in the collection.</returns>
        /// <exception cref="T:System.InvalidOperationException">The enumerator is positioned before the first element of the collection or after the last element.-or- The collection was modified after the enumerator was created.</exception>
        object IEnumerator.Current
        {
            get
            {
                this.lines.SetCurrentLine(this.currentLine);
                return this.lines;
            }
        }

        /// <summary>
        /// Gets the current element in the collection.
        /// </summary>
        /// <value></value>
        /// <returns>The current element in the collection.</returns>
        /// <exception cref="T:System.InvalidOperationException">The enumerator is positioned before the first element of the collection or after the last element.-or- The collection was modified after the enumerator was created.</exception>
        public Document_Lines Current
        {
            get
            {
                this.lines.SetCurrentLine(this.currentLine);
                return this.lines;
            }
        }

        /// <summary>
        /// Advances the enumerator to the next element of the collection.
        /// </summary>
        /// <returns>
        /// true if the enumerator was successfully advanced to the next element; false if the enumerator has passed the end of the collection.
        /// </returns>
        /// <exception cref="T:System.InvalidOperationException">The collection was modified after the enumerator was created. </exception>
        public bool MoveNext()
        {
            if (this.currentLine < this.lines.Count)
            {
                this.currentLine++;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Sets the enumerator to its initial position, which is before the first element in the collection.
        /// </summary>
        /// <exception cref="T:System.InvalidOperationException">The collection was modified after the enumerator was created. </exception>
        public void Reset()
        {
            this.currentLine = 0;
        }

        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected override void Release()
        {
            COMHelper.Release(ref this.lines);
        }

        #endregion

        #region IEnumerable Members
        /// <summary>
        /// Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>
        /// An <see cref="T:System.Collections.IEnumerator"/> object that can be used to iterate through the collection.
        /// </returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            int lineCount = this.lines.Count;

            for (int lineIndex = 0; lineIndex < lineCount; lineIndex++)
            {
                this.currentLine = lineIndex;
                yield return this.Current;
            }
        }
        #endregion IEnumerable Members

        /// <summary>
        /// Adds the specified line num.
        /// </summary>
        /// <param name="lineNum">The line num.</param>
        /// <returns>The SAP lines added.</returns>
        public Document_Lines Add(int lineNum)
        {
            if (lineNum >= this.lines.Count)
            {
                this.lines.Add();
            }

            this.lines.SetCurrentLine(lineNum);
            this.currentLine = lineNum;
            return this.lines;
        }

        /// <summary>
        /// Sets the user defined field on the current line.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <param name="fieldValue">The field value.</param>
        public void SetUserDefinedField(object fieldIndex, object fieldValue)
        {
            UserFields uf = this.Current.UserFields;
            try
            {
                Fields fields = uf.Fields;
                try
                {
                    Field field = fields.Item(fieldIndex);
                    try
                    {
                        field.Value = fieldValue;
                    }
                    finally
                    {
                        COMHelper.Release(ref field);
                    }
                }
                finally
                {
                    COMHelper.Release(ref fields);
                }
            }
            finally
            {
                COMHelper.Release(ref uf);
            }
        }

        public void SetCurrentLine(int lineIndex)
        {
            this.currentLine = lineIndex;
            this.lines.SetCurrentLine(lineIndex);
        }

        /// <summary>
        /// Gets the user defined field from the current line.
        /// </summary>
        /// <param name="fieldIndex">Index of the field.</param>
        /// <returns>The value of the field</returns>
        public object GetUserDefinedField(object fieldIndex)
        {
            object fieldValue = null;

            UserFields uf = this.Current.UserFields;
            try
            {
                Fields fields = uf.Fields;
                try
                {
                    Field field = fields.Item(fieldIndex);
                    try
                    {
                        fieldValue = field.Value;
                    }
                    finally
                    {
                        COMHelper.Release(ref field);
                    }
                }
                finally
                {
                    COMHelper.Release(ref fields);
                }
            }
            finally
            {
                COMHelper.Release(ref uf);
            }

            return fieldValue;
        }

        /// <summary>
        /// Copies the current line UDFs to the given line's current line
        /// </summary>
        /// <param name="line">The line to receive the UDF values.</param>
        public void CopyCurrentLineUDFsTo(Lines line)
        {
            UserFields uf = this.Current.UserFields;
            UserFields toUf = line.Current.UserFields;
            try
            {
                Fields fields = uf.Fields;
                Fields toFields = toUf.Fields;
                try
                {
                    int fieldCount = fields.Count;
                    for (int i = 0; i < fieldCount; i++)
                    {
                        Field field = fields.Item(i);
                        Field toField = toFields.Item(i);

                        try
                        {
                            toField.Value = field.Value;
                        }
                        finally
                        {
                            COMHelper.Release(ref field);
                            COMHelper.Release(ref toField);
                        }
                    }
                }
                finally
                {
                    COMHelper.Release(ref fields);
                    COMHelper.Release(ref toFields);
                }
            }
            finally
            {
                COMHelper.Release(ref uf);
                COMHelper.Release(ref toUf);
            }
        }

        #endregion Methods
    }
}
