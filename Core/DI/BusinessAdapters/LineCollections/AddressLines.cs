//-----------------------------------------------------------------------
// <copyright file="AddressLines.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
using B1C.SAP.DI.Helpers;

namespace B1C.SAP.DI.BusinessAdapters
{
    using System;
    using System.Collections;
    
    /// <summary>
    /// Instance of the AddressLines class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public class AddressLines : SapBaseObject, IEnumerable, IEnumerator
    {
        #region Fields
        /// <summary>
        /// Business Partner lines.
        /// </summary>
        private SAPbobsCOM.BPAddresses lines;

        /// <summary>
        /// Index of the current line.
        /// </summary>
        private int currentLine;
        #endregion Fields

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="AddressLines"/> class.
        /// </summary>
        /// <param name="lines">The Business Partner lines.</param>
        public AddressLines(SAPbobsCOM.BPAddresses lines)
        {
            this.lines = lines;
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
        public SAPbobsCOM.BPAddresses Current
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
        /// <param name="lineNum">The line number.</param>
        /// <returns>The address added.</returns>
        public SAPbobsCOM.BPAddresses Add(int lineNum)
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
        /// Releases the Documents COM Object.
        /// </summary>
        protected override void Release()
        {
            COMHelper.Release(ref this.lines);
        }
        #endregion Methods
    }
}
