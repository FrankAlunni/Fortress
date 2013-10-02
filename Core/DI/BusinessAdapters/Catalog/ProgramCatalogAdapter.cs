// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ProgramCatalogAdapter.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the Program Catalog type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.Catalog
{
    #region Usings
    using System;
    using System.Text;
    using Helpers;
    using SAPbobsCOM;
    using Utility.Helpers;
    #endregion Usings

    /// <summary>
    /// Instance of a Program catalog class
    /// </summary>
    public class ProgramCatalogAdapter : ProgramCatalogObjectAdapter
    {
        #region Fields 
        /// <summary>
        /// Then name of the catalog table
        /// </summary>
        private const string MainTableName = "@XX_CTPM";
        #endregion Fields

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="ProgramCatalogAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="programCode">The program code.</param>
        public ProgramCatalogAdapter(Company company, string programCode) : base(company)
        {
            var sql = new SqlHelper();
            sql.Builder.AppendFormat("SELECT * FROM [{0}] WHERE U_CardCode = @CardCode AND ISNULL(U_Active, 'N') = @Active", MainTableName);
            sql.AddParameter("@CardCode", System.Data.DbType.String, programCode);
            sql.AddParameter("@Active", System.Data.DbType.String, 'Y');
            using (var recordSet = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                if (recordSet.RecordCount > 1)
                {
                    throw new Exception(string.Format("The program {0} has more than one active catalog.", programCode));
                }

                if (recordSet.RecordCount == 0)
                {
                    throw new Exception(string.Format("The program {0} does not have any active catalog.", programCode));
                }

                this.Code = recordSet.FieldValue("Code").ToString();
                this.Name = recordSet.FieldValue("Name").ToString();
                this.CardCode = recordSet.FieldValue("U_CardCode").ToString();
                this.GlobalCatalogCode = recordSet.FieldValue("U_GlbCatCd").ToString();
                this.Description = recordSet.FieldValue("U_Descript").ToString();
                this.OverrideImage = recordSet.FieldValue("U_OverrSz").ToString();

                int size;
                int.TryParse(recordSet.FieldValue("U_ThumbSz").ToString(), out size);
                this.OverrideThumbnailSize = size;

                int.TryParse(recordSet.FieldValue("U_SmallSz").ToString(), out size);
                this.OverrideSmallImageSize = size;

                int.TryParse(recordSet.FieldValue("U_LargeSz").ToString(), out size);
                this.OverrideLargeImageSize = size;

                this.Active = true;

                DateTime databaseDate;
                DateTime.TryParse(recordSet.FieldValue("U_StartDt").ToString(), out databaseDate);
                this.StartDate = databaseDate;

                DateTime.TryParse(recordSet.FieldValue("U_EndDt").ToString(), out databaseDate);
                this.EndDate = databaseDate;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ProgramCatalogAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="programCode">The program code.</param>
        /// <param name="catalogCode">The catalog code.</param>
        public ProgramCatalogAdapter(Company company, string programCode, string catalogCode) : base(company)
        {
            var sql = new SqlHelper();
            sql.Builder.AppendFormat("SELECT * FROM [{0}] WHERE Code = @Code AND U_CardCode = @CardCode", MainTableName);
            sql.AddParameter("@Code", System.Data.DbType.String, catalogCode);
            sql.AddParameter("@CardCode", System.Data.DbType.String, programCode);
            using (var recordSet = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                if (recordSet.RecordCount == 0)
                {
                    throw new Exception(string.Format("The catalog {0} is not found or is not for program {1}.", catalogCode, programCode));
                }

                this.Code = recordSet.FieldValue("Code").ToString();
                this.Name = recordSet.FieldValue("Name").ToString();
                this.CardCode = recordSet.FieldValue("U_CardCode").ToString();
                this.GlobalCatalogCode = recordSet.FieldValue("U_GlbCatCd").ToString();
                this.Description = recordSet.FieldValue("U_Descript").ToString();
                this.OverrideImage = recordSet.FieldValue("U_OverrSz").ToString();

                int size;
                int.TryParse(recordSet.FieldValue("U_ThumbSz").ToString(), out size);
                this.OverrideThumbnailSize = size;

                int.TryParse(recordSet.FieldValue("U_SmallSz").ToString(), out size);
                this.OverrideSmallImageSize = size;

                int.TryParse(recordSet.FieldValue("U_LargeSz").ToString(), out size);
                this.OverrideLargeImageSize = size;

                this.Active = recordSet.FieldValue("U_Descript").ToString() == "Y" ? true : false;

                DateTime databaseDate;
                DateTime.TryParse(recordSet.FieldValue("U_StartDt").ToString(), out databaseDate);
                this.StartDate = databaseDate;

                DateTime.TryParse(recordSet.FieldValue("U_EndDt").ToString(), out databaseDate);
                this.EndDate = databaseDate;
            }
        }
        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets or sets the override image.
        /// </summary>
        /// <value>The override image.</value>
        public string OverrideImage { get; set; }

        /// <summary>
        /// Gets or sets the size of the override thumbnail.
        /// </summary>
        /// <value>The size of the override thumbnail.</value>
        public int OverrideThumbnailSize { get; set; }

        /// <summary>
        /// Gets or sets the size of the override small image.
        /// </summary>
        /// <value>The size of the override small image.</value>
        public int OverrideSmallImageSize { get; set; }

        /// <summary>
        /// Gets or sets the size of the override large image.
        /// </summary>
        /// <value>The size of the override large image.</value>
        public int OverrideLargeImageSize { get; set; }
        #endregion Properties

        #region Methods

        /// <summary>
        /// Cleans this instance.
        /// </summary>
        /// <param name="shipValidation">if set to <c>true</c> [ship validation].</param>
        /// <returns>True if the method runs successfully</returns>
        public bool Clean(bool shipValidation)
        {
            if (!shipValidation)
            {
                if (this.Active)
                {
                    throw new Exception(string.Format("The catalog {0} is active. An active catalog cannot be cleaned.", this.Code));
                }

                if (this.StartDate < DateTime.Now)
                {
                    throw new Exception(string.Format("The catalog {0} has already started. A started catalog cannot be cleaned.", this.Code));
                }
            }

            var builder = new StringBuilder();
            builder.AppendLine("DECLARE  @Error int; SET @Error = 1");
            builder.AppendLine("BEGIN TRANSACTION");
            builder.AppendLine("--Delete All The Items in Inactive Categories");
            builder.AppendLine("DELETE FROM [@XX_CTP1] WHERE Code IN ");
            builder.AppendFormat("(SELECT CategoryCode FROM dbo.oz_CategoriesInactive('{0}', ''))", this.Code);
            builder.AppendLine(string.Empty);
            builder.AppendLine("IF (@@Error <> 0) GOTO ENDPROCESS");
            builder.AppendLine("--Delete All Inactive Items In the Active Categories");
            builder.AppendLine("DELETE FROM [@XX_CTP1] WHERE ISNULL(U_Active, 'N') = 'N'");
            builder.AppendFormat("AND Code IN (SELECT CategoryCode FROM dbo.oz_CategoriesActive('{0}', ''))", this.Code);
            builder.AppendLine(string.Empty);
            builder.AppendLine("IF (@@Error <> 0) GOTO ENDPROCESS");
            builder.AppendLine("--Delete All Inactive Categories");
            builder.AppendLine("DELETE FROM [@XX_CTPC] WHERE Code IN ");
            builder.AppendFormat("(SELECT CategoryCode FROM dbo.oz_CategoriesInactive('{0}', ''))", this.Code);
            builder.AppendLine(string.Empty);
            builder.AppendLine("IF (@@Error <> 0) GOTO ENDPROCESS");
            builder.AppendLine("-- Delete Items Not In Any Catalog");
            builder.AppendLine("DELETE FROM [@XX_CTP1] WHERE CODE IN  (");
            builder.AppendLine("SELECT Code FROM [@XX_CTPC] WHERE ISNULL(U_CatCode, '') = '')");
            builder.AppendLine("IF (@@Error <> 0) GOTO ENDPROCESS");
            builder.AppendLine("--Delete Categories not in any Catalog");
            builder.AppendLine("DELETE FROM [@XX_CTPC] WHERE ISNULL(U_CatCode, '') = ''");
            builder.AppendLine("IF (@@Error <> 0) GOTO ENDPROCESS");
            builder.AppendLine("SET @Error = 0");
            builder.AppendLine("ENDPROCESS:");
            builder.AppendLine("IF @Error = 0");
            builder.AppendLine("BEGIN");
            builder.AppendLine("	COMMIT TRANSACTION");
            builder.AppendLine("END");
            builder.AppendLine("ELSE");
            builder.AppendLine("BEGIN");
            builder.AppendLine("	ROLLBACK TRANSACTION");
            builder.AppendLine("END");

            RecordsetAdapter.Execute(this.Company, builder.ToString());

            return true;
        }

        /// <summary>
        /// Cleans this instance.
        /// </summary>
        /// <returns>True if the method runs successfully</returns>
        public bool Clean()
        {
            return this.Clean(false);
        }

        /// <summary>
        /// Copies the translations from item master.
        /// </summary>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        /// <returns>True if the translation is copied successfully</returns>
        public bool CopyTranslationsFromItemMaster(bool overwrite)
        {
            var catalogRecordset = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            catalogRecordset.DoQuery(string.Format("select Code from [@XX_CTPM] where U_CardCode = '{0}' and U_Active = 'Y'", this.CardCode));
            var catalogRecordsetHelper = COMHelper.RecordSet(catalogRecordset);

            try
            {
                if (!catalogRecordset.EoF)
                {
                    string catalogCode = catalogRecordsetHelper.FieldValue("Code").ToString();

                    var categoryRecordset = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    categoryRecordset.DoQuery(string.Format("select CategoryCode from dbo.oz_CategoriesAll('{0}', '')", catalogCode));
                    var categoryRecordsetHelper = COMHelper.RecordSet(categoryRecordset);

                    try
                    {
                        while (!categoryRecordset.EoF)
                        {
                            string categoryCode = categoryRecordsetHelper.FieldValue("CategoryCode").ToString();

                            var categoryItemRecordset = (Recordset)this.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                            string itemQuery = string.Format("select (Code + Cast(LineId as varchar(10))) as PK, U_ItemCode from [@XX_CTP1] where Code = '{0}'", categoryCode);
                            categoryItemRecordset.DoQuery(itemQuery);
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
                                            if (itemNameTranRecordset.EoF)
                                            {
                                                new TableFieldAdapter(this.Company, "OITM", "ItemName").MoveTranslation(itemCode, "@XX_CTP1", "U_ItemName", translationKey);
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
                                            if (descriptionTranRecordset.EoF)
                                            {
                                                new TableFieldAdapter(this.Company, "OITM", "U_LongDesc").MoveTranslation(itemCode, "@XX_CTP1", "U_Descript", translationKey);
                                            }
                                        }
                                        finally
                                        {
                                            descriptionTranRecordsetHelper.Release();
                                        }
                                    }
                                    else
                                    {
                                        // Copy the Item Name
                                        new TableFieldAdapter(this.Company, "OITM", "ItemName").MoveTranslation(itemCode, "@XX_CTP1", "U_ItemName", translationKey);

                                        // Copy the Item Description
                                        new TableFieldAdapter(this.Company, "OITM", "U_LongDesc").MoveTranslation(itemCode, "@XX_CTP1", "U_Descript", translationKey);
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

        #endregion Methods
    }
}
