//-----------------------------------------------------------------------
// <copyright file="Toolbar.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.UI.Constants
{
    using System.Collections.Specialized;
    /// <summary>
    /// The tool bar
    /// </summary>
    public class Toolbar
    {
        public const string Save = "516";
        public const string PrintPreview = "519";
        public const string Print = "520";
        public const string SendEmail = "6657";
        public const string SendSms = "6658";
        public const string SendFax = "6659";
        public const string ExportToMsExcel = "7169";
        public const string ExportToMsWord = "7170";
        public const string LaunchApplication = "523";
        public const string LockScreen = "524";
        public const string Find = "1281";
        public const string Add = "1282";
        public const string FirstDataRecord = "1290";
        public const string PreviousRecord = "1289";
        public const string NextRecord = "1288";
        public const string LastDataRecord = "1291";
        public const string DocumentEditing = "5895";
        public const string TransactionJournal = "5894";
        public const string PaymentMeans = "5892";
        public const string GrossProfit = "5891";
        public const string VolumeWeightCalculation = "5893";
        public const string BaseDocument = "5898";
        public const string TargetDocument = "5899";
        public const string Settings = "5890";
        public const string SortTable = "4869";
        public const string SavedQueries = "4865";
        public const string DeleteRow = "1293";
        public const string DuplicateRow = "1294";

        /// <summary>
        /// The toolbar items
        /// </summary>
        private readonly StringDictionary _toobarItems;

        /// <summary>
        /// Initializes a new instance of the <see cref="Toolbar"/> class.
        /// </summary>
        public Toolbar()
        {
            _toobarItems = new StringDictionary
                              {
                                  {"516", "Save"},
                                  {"519", "PrintPreview"},
                                  {"520", "Print"},
                                  {"6657", "SendEmail"},
                                  {"6658", "SendSms"},
                                  {"6659", "SendFax"},
                                  {"7169", "ExportToMsExcel"},
                                  {"7170", "ExportToMsWord"},
                                  {"523", "LaunchApplication"},
                                  {"524", "LockScreen"},
                                  {"1281", "Find"},
                                  {"1282", "Add"},
                                  {"1290", "FirstDataRecord"},
                                  {"1289", "PreviousRecord"},
                                  {"1288", "NextRecord"},
                                  {"1291", "LastDataRecord"},
                                  {"5895", "DocumentEditing"},
                                  {"5894", "TransactionJournal"},
                                  {"5892", "PaymentMeans"},
                                  {"5891", "GrossProfit"},
                                  {"5893", "VolumeWeightCalculation"},
                                  {"5898", "BaseDocument"},
                                  {"5899", "TargetDocument"},
                                  {"5890", "Settings"},
                                  {"4869", "SortTable"},
                                  {"4865", "SavedQueries"},
                                  {"1293", "DeleteRow"},
                                  {"1294", "DuplicateRow"}

                              };
        }

        /// <summary>
        /// Determines whether [is tool bar item] [the specified item id].
        /// </summary>
        /// <param name="itemId">The item id.</param>
        /// <returns>
        /// 	<c>true</c> if [is tool bar item] [the specified item id]; otherwise, <c>false</c>.
        /// </returns>
        public bool ContainsItem(string itemId)
        {
            return this._toobarItems.ContainsKey(itemId);
        }
    }
}
