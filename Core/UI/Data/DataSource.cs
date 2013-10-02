// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DataSource.cs" company="B1C Canada Inc.">
//   B1C Canada Inc.
// </copyright>
// <summary>
//   Defines the DataSource type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.UI.Data
{
    using Windows;

    /// <summary>
    /// Form Data sources
    /// </summary>
    public class DataSource
    {
        /// <summary>
        /// AddOn Instance
        /// </summary>
        private Form form;

        /// <summary>
        /// Initializes a new instance of the <see cref="DataSource"/> class.
        /// </summary>
        /// <param name="currentForm">The current form.</param>
        public DataSource(Form currentForm)
        {
            this.form = currentForm;
        }

        /// <summary>
        /// Databases the specified source name.
        /// </summary>
        /// <param name="sourceName">Name of the source.</param>
        /// <returns>The requested datasource</returns>
        public SAPbouiCOM.DBDataSource SystemTable(string sourceName)
        {
            return this.form.SapForm.DataSources.DBDataSources.Item(sourceName);
        }

        /// <summary>
        /// Datas the table.
        /// </summary>
        /// <param name="sourceName">Name of the source.</param>
        /// <returns>The requested datasource</returns>
        public SAPbouiCOM.DataTable UserTable(string sourceName)
        {
            return this.form.SapForm.DataSources.DataTables.Item(sourceName);
        }

        /// <summary>
        /// Users the source.
        /// </summary>
        /// <param name="sourceName">Name of the source.</param>
        /// <returns>The requested datasource</returns>
        public SAPbouiCOM.UserDataSource UserSource(string sourceName)
        {
            return this.form.SapForm.DataSources.UserDataSources.Item(sourceName);
        }
    }
}
