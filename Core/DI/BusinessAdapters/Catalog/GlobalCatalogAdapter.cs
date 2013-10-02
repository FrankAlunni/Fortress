// --------------------------------------------------------------------------------------------------------------------
// <copyright file="GlobalCatalogAdapter.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the GlobalCatalogAdapter type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using B1C.SAP.DI.Helpers;

namespace B1C.SAP.DI.BusinessAdapters.Catalog
{
    using SAPbobsCOM;

    /// <summary>
    /// The global catalog adapter
    /// </summary>
    public class GlobalCatalogAdapter : SapObjectAdapter
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="GlobalCatalogAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        public GlobalCatalogAdapter(Company company)
            : base(company)
        {
        }

        #endregion Constructors

        #region Method(s)

        /// <summary>
        /// Copies the translations from item master.
        /// </summary>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        /// <returns>True if the translation is copied successfully</returns>
        public bool CopyTranslationsFromItemMaster(bool overwrite)
        {
            var catalogRecordset = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            catalogRecordset.DoQuery(string.Format("select Code from [@XX_CTCM] where U_Active = 'Y'"));
            var catalogRecordsetHelper = COMHelper.RecordSet(catalogRecordset);

            try
            {
                if (!catalogRecordset.EoF)
                {
                    string catalogCode = catalogRecordsetHelper.FieldValue("Code").ToString();

                    var categoryRecordset = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    categoryRecordset.DoQuery(string.Format("select Code from [@XX_CTCC] where U_CatCode = '{0}'", catalogCode));
                    var categoryRecordsetHelper = COMHelper.RecordSet(categoryRecordset);

                    try
                    {
                        while (!categoryRecordset.EoF)
                        {
                            string categoryCode = categoryRecordsetHelper.FieldValue("Code").ToString();

                            var categoryItemRecordset = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                            categoryItemRecordset.DoQuery(string.Format("select (Code + Cast(LineId as varchar(10))) as PK, U_ItemCode from [@XX_CTC1] where Code = '{0}'", categoryCode));
                            var categoryItemRecordsetHelper = COMHelper.RecordSet(categoryItemRecordset);

                            try
                            {
                                while (!categoryItemRecordset.EoF)
                                {
                                    string itemCode = categoryItemRecordsetHelper.FieldValue("U_ItemCode").ToString();
                                    string translationKey = categoryItemRecordsetHelper.FieldValue("PK").ToString();

                                    if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(translationKey))
                                    {
                                        categoryItemRecordset.MoveNext();
                                        continue;
                                    }

                                    // If we're not overwriting we need to check if translations already exist
                                    if (!overwrite)
                                    {
                                        string itemNameTranslationExistsQuery = string.Format("select 1 from OMLT where PK = '{0}' and FieldAlias = 'U_ItemName' and TableName = '@XX_CTP1'", translationKey);
                                        var itemNameTranRecordset = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        itemNameTranRecordset.DoQuery(itemNameTranslationExistsQuery);
                                        var itemNameTranRecordsetHelper = COMHelper.RecordSet(itemNameTranRecordset);

                                        try
                                        {
                                            if (!itemNameTranRecordset.EoF)
                                            {
                                                new TableFieldAdapter(this.Company, "OITM", "ItemName").MoveTranslation(itemCode, "@XX_CTC1", "U_ItemName", translationKey);
                                            }
                                        }
                                        finally
                                        {
                                            itemNameTranRecordsetHelper.Release();                                            
                                        }

                                        string descriptionTranslationExistsQuery = string.Format("select 1 from OMLT where PK = '{0}' and FieldAlias = 'U_Descript' and TableName = '@XX_CTP1'", translationKey);
                                        var descriptionTranRecordset = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        descriptionTranRecordset.DoQuery(descriptionTranslationExistsQuery);
                                        var descriptionTranRecordsetHelper = COMHelper.RecordSet(descriptionTranRecordset);

                                        try
                                        {
                                            if (!descriptionTranRecordset.EoF)
                                            {
                                                new TableFieldAdapter(this.Company, "OITM", "U_LongDesc").MoveTranslation(itemCode, "@XX_CTC1", "U_Descript", translationKey);
                                            }
                                        }
                                        finally
                                        {
                                            descriptionTranRecordsetHelper.Release();    
                                        }
                                    }

                                    categoryItemRecordset.MoveNext();
                                }
                            }
                            finally
                            {
                                categoryItemRecordsetHelper.Release();                                
                            }
                            
                            categoryRecordset.MoveNext();
                        }
                    }
                    finally
                    {
                        categoryRecordsetHelper.Release();                        
                    }
                }
            }
            finally
            {
                catalogRecordsetHelper.Release();
            }

            return true;
        }

        /// <summary>
        /// Releases the COM Objects.
        /// </summary>
        protected override void Release()
        {
        }
        #endregion Method(s)

    }
}
