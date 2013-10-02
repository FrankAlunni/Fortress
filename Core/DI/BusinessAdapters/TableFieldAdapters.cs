// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TableFieldAdapters.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the TableFields type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters
{
    using System.Collections;
    using SAPbobsCOM;
    using System.Collections.Generic;

    /// <summary>
    /// Table Field collection
    /// </summary>
    public class TableFieldAdapters : SapBaseObject, IEnumerable, IEnumerator
    {
        private List<TableFieldAdapter> list;
        private int current = -1;
        
        public TableFieldAdapters()
        {
            list = new List<TableFieldAdapter>();
        }

        /// <summary>
        /// Gets the <see cref="TableFieldAdapter"/> at the specified index.
        /// </summary>
        /// <param name="index">The index in the collection</param>
        /// <value>The table field class</value>
        public virtual TableFieldAdapter this[int index]
        {
            get
            {
                return list[index];
            }
        }

        /// <summary>
        /// Adds the specified company.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="tableField">The table field.</param>
        public virtual TableFieldAdapter Add(Company company, TableFieldAdapter tableField)
        {
            list.Add(tableField);
            return list[list.Count-1];
        }

        /// <summary>
        /// Adds the specified company.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="fieldDescription">The field description.</param>
        /// <param name="fieldSize">Size of the field.</param>
        /// <param name="fieldEditSize">Size of the field edit.</param>
        /// <param name="fieldType">Type of the field.</param>
        /// <param name="fieldSubType">Type of the field sub.</param>
        /// <param name="linkedTable">The linked table.</param>
        public virtual TableFieldAdapter Add(
            Company company,
            string tableName,
            string fieldName,
            string fieldDescription,
            int? fieldSize,
            int? fieldEditSize,
            BoFieldTypes fieldType,
            BoFldSubTypes? fieldSubType,
            string linkedTable)
        {
            var field = new TableFieldAdapter(company, tableName, fieldName)
                            {
                                TableName = tableName,
                                Name = fieldName,
                                Description = fieldDescription,
                                Size = fieldSize,
                                EditSize = fieldEditSize,
                                Type = fieldType,
                                SubType = fieldSubType,
                                LinkedTable = linkedTable
                            };

            return this.Add(company, field);
        }

        /// <summary>
        /// Adds the specified company.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="fieldDescription">The field description.</param>
        /// <param name="fieldSize">Size of the field.</param>
        /// <param name="fieldEditSize">Size of the field edit.</param>
        /// <param name="fieldType">Type of the field.</param>
        /// <param name="fieldSubType">Type of the field sub.</param>
        /// <param name="linkedTable">The linked table.</param>
        /// <param name="validValues">The valid values.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns></returns>
        public virtual TableFieldAdapter Add(
            Company company,
            string tableName,
            string fieldName,
            string fieldDescription,
            int? fieldSize,
            int? fieldEditSize,
            BoFieldTypes fieldType,
            BoFldSubTypes? fieldSubType,
            string linkedTable,
            Dictionary<string, string> validValues,
            string defaultValue)
        {
            var field = new TableFieldAdapter(company, tableName, fieldName)
                            {
                                TableName = tableName,
                                Name = fieldName,
                                Description = fieldDescription,
                                Size = fieldSize,
                                EditSize = fieldEditSize,
                                Type = fieldType,
                                SubType = fieldSubType,
                                LinkedTable = linkedTable,
                                ValidValues = validValues,
                                DefaultValue = defaultValue
                            };


            return this.Add(company, field);
        }

        /// <summary>
        /// Releases this instance.
        /// </summary>
        protected override void Release()
        {
            for (var i = this.list.Count - 1; i >= 0; i--)
            {
                this.list[i].Dispose();
                this.list[i] = null;                
            }
        }

        /// <summary>
        /// Gets the count.
        /// </summary>
        /// <value>The count.</value>
        public int Count
        {
            get
            {
                if (this.list == null)
                {
                    return 0;
                }

                return this.list.Count;
            }
        }

        #region IEnumerable Members

        public IEnumerator GetEnumerator()
        {
            int lineCount = this.list.Count;

            for (int lineIndex = 0; lineIndex < lineCount; lineIndex++)
            {
                this.current = lineIndex;
                yield return this.Current;
            }
        }

        #endregion

        #region IEnumerator Members

        public object Current
        {
            get
            {
                if ((current == -1) || (current >= this.list.Count))
                {
                    return null;
                }

                return list[current];
            }
        }

        public bool MoveNext()
        {
            if (this.current < this.list.Count)
            {
                this.current++;
                return true;
            }

            return false;
        }

        public void Reset()
        {
            current = -1;
        }

        #endregion
    }
}
