// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ProgramCategoryAdapter.cs" company="B1C Canada Inc.">
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
    using System.Collections.Generic;
    using SAPbobsCOM;
    using Utility.Helpers;
    #endregion Usings

    /// <summary>
    /// Instance of a Program catalog class
    /// </summary>
    public class ProgramCategoryAdapter : ProgramCatalogObjectAdapter
    {
        #region Fields 
        /// <summary>
        /// Then name of the catalog table
        /// </summary>
        private const string MainTableName = "@XX_CTPC";

        /// <summary>
        /// The active items not in pricing.
        /// </summary>
        private IList<string> activeItemCodesNotInPricing;

        /// <summary>
        /// The inactive items in pricing.
        /// </summary>
        private IList<string> inactiveItemCodesInPricing;

        /// <summary>
        /// Active Items in the category
        /// </summary>
        private IList<string> activeItemCodes;

        /// <summary>
        /// Inctive Items in the category.
        /// </summary>
        private IList<string> inactiveItemCodes;

        /// <summary>
        /// A List of active category codes
        /// </summary>
        private IList<string> activeCategoryCodes;

        /// <summary>
        /// A List of inactive category codes
        /// </summary>
        private IList<string> inactiveCategoryCodes;
        #endregion Fields

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="ProgramCategoryAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="categoryCode">The category code.</param>
        /// <param name="cardCode">The card code.</param>
        public ProgramCategoryAdapter(Company company, string categoryCode, string cardCode) : base(company)
        {
            var sql = new SqlHelper();
            sql.Builder.AppendFormat("SELECT * FROM [{0}] WHERE Code = @Code", MainTableName);
            sql.AddParameter("@Code", System.Data.DbType.String, categoryCode);
            using (var recordSet = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                if (recordSet.RecordCount == 0)
                {
                    throw new Exception(string.Format("The category {0} does not exists.", categoryCode));
                }

                this.Code = recordSet.FieldValue("Code").ToString();
                this.Name = recordSet.FieldValue("Name").ToString();
                this.CatalogCode = recordSet.FieldValue("U_CatCode").ToString();
                this.ParentCode = recordSet.FieldValue("U_Parent").ToString();
                this.GlobalCatalogCode = recordSet.FieldValue("U_GlbCatCd").ToString();
                this.Description = recordSet.FieldValue("U_Descript").ToString();
                this.Active = recordSet.FieldValue("U_LargeSz").ToString() == "Y";
                DateTime databaseDate;
                DateTime.TryParse(recordSet.FieldValue("U_StartDt").ToString(), out databaseDate);
                this.StartDate = databaseDate;

                DateTime.TryParse(recordSet.FieldValue("U_EndDt").ToString(), out databaseDate);
                this.EndDate = databaseDate;
                this.CardCode = cardCode;
            }
        }
        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets or sets the catalog code.
        /// </summary>
        /// <value>The catalog code.</value>
        public string CatalogCode { get; set; }

        /// <summary>
        /// Gets or sets the parent code.
        /// </summary>
        /// <value>The parent code.</value>
        public string ParentCode { get; set; }

        /// <summary>
        /// Gets the inactive item codes in pricing.
        /// </summary>
        /// <value>The inactive item codes in pricing.</value>
        public override IList<string> InactiveItemCodesInPricing
        {
            get
            {
                if (this.inactiveItemCodesInPricing == null)
                {
                    this.inactiveItemCodesInPricing = this.GetInactiveItemsInPricing();
                }

                return this.inactiveItemCodesInPricing;
            }
        }

        /// <summary>
        /// Gets the active items not in pricing.
        /// </summary>
        /// <value>The active items not in pricing.</value>
        public override IList<string> ActiveItemCodesNotInPricing
        {
            get
            {
                if (this.activeItemCodesNotInPricing == null)
                {
                    this.activeItemCodesNotInPricing = this.GetActiveItemsNotInPricing();
                }

                return this.activeItemCodesNotInPricing;
            }
        }

        /// <summary>
        /// Gets the inactive item codes in pricing.
        /// </summary>
        /// <value>The inactive item codes in pricing.</value>
        public override IList<string> InactiveItemCodes
        {
            get
            {
                if (this.inactiveItemCodes == null)
                {
                    this.inactiveItemCodes = this.GetInactiveItems();
                }

                return this.inactiveItemCodes;
            }
        }

        /// <summary>
        /// Gets the active items not in pricing.
        /// </summary>
        /// <value>The active items not in pricing.</value>
        public override IList<string> ActiveItemCodes
        {
            get
            {
                if (this.activeItemCodes == null)
                {
                    this.activeItemCodes = this.GetActiveItems();
                }

                return this.activeItemCodes;
            }
        }

        /// <summary>
        /// Gets the active items not in pricing.
        /// </summary>
        /// <value>The active items not in pricing.</value>
        public override IList<string> ActiveCategoryCodes
        {
            get
            {
                if (this.activeCategoryCodes == null)
                {
                    this.activeCategoryCodes = this.GetActiveCategories();
                }

                return this.activeCategoryCodes;
            }
        }

        /// <summary>
        /// Gets the inactive category codes.
        /// </summary>
        /// <value>The inactive category codes.</value>
        public override IList<string> InactiveCategoryCodes
        {
            get
            {
                if (this.inactiveCategoryCodes == null)
                {
                    this.inactiveCategoryCodes = this.GetInactiveCategories();
                }

                return this.inactiveCategoryCodes;
            }
        } 
        #endregion Properties

        #region Methods
        /// <summary>
        /// Gets the active items not in pricing.
        /// </summary>
        /// <returns>The list of items not in the pricing but in the catalog</returns>
        private IList<string> GetActiveItemsNotInPricing()
        {
            IList<string> itemCodes = new List<string>();
            var sql = new SqlHelper();
            sql.Builder.AppendLine("SELECT T0.ItemCode");
            sql.Builder.AppendLine("FROM dbo.oz_ItemCategoryActive(@Code) T0");
            sql.Builder.AppendLine("LEFT JOIN OSCN T1 ON T1.Code = @Code AND T0.ItemCode = T1.ItemCode");
            sql.Builder.AppendLine("WHERE T1.ItemCode  IS NULL");
            sql.AddParameter("@Code", System.Data.DbType.String, this.Code);

            using (var recordset = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                while (!recordset.EoF)
                {
                    itemCodes.Add(recordset.FieldValue("ItemCode").ToString());
                    recordset.MoveNext();
                }
            }

            return itemCodes;
        }

        /// <summary>
        /// Gets the inactive items in pricing.
        /// </summary>
        /// <returns>The list of items not active but in the pricing </returns>
        private IList<string> GetInactiveItemsInPricing()
        {
            IList<string> itemCodes = new List<string>();
            var sql = new SqlHelper();
            sql.Builder.AppendLine("SELECT T0.ItemCode");
            sql.Builder.AppendLine("FROM OSCN T0");
            sql.Builder.AppendLine("LEFT JOIN dbo.oz_ItemCategoryActive(@Code) T1");
            sql.Builder.AppendLine("ON T0.Code = @Code AND T0.ItemCode = T1.ItemCode");
            sql.Builder.AppendLine("WHERE T1.ItemCode  IS NULL AND T0.Code = @Code");
            sql.AddParameter("@Code", System.Data.DbType.String, this.Code);

            using (var recordset = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                while (!recordset.EoF)
                {
                    itemCodes.Add(recordset.FieldValue("ItemCode").ToString());
                    recordset.MoveNext();
                }
            }

            return itemCodes;
        }

        /// <summary>
        /// Gets the active items not in pricing.
        /// </summary>
        /// <returns>The list of items not in the pricing but in the catalog</returns>
        private IList<string> GetActiveItems()
        {
            IList<string> itemCodes = new List<string>();
            var sql = new SqlHelper();
            sql.Builder.AppendLine("SELECT U_ItemsCode AS ItemCode FROM [@XX_CTP1] WHERE Code = @Code AND ISNULL(U_Active, 'N') = 'Y'");
            sql.AddParameter("@Code", System.Data.DbType.String, this.Code);

            using (var recordset = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                while (!recordset.EoF)
                {
                    itemCodes.Add(recordset.FieldValue("ItemCode").ToString());
                    recordset.MoveNext();
                }
            }

            return itemCodes;
        }

        /// <summary>
        /// Gets the inactive items in pricing.
        /// </summary>
        /// <returns>The list of items not active but in the pricing </returns>
        private IList<string> GetInactiveItems()
        {
            IList<string> itemCodes = new List<string>();
            var sql = new SqlHelper();
            sql.Builder.AppendLine("SELECT U_ItemsCode AS ItemCode FROM [@XX_CTP1] WHERE Code = @Code AND ISNULL(U_Active, 'N') = 'Y'");
            sql.AddParameter("@Code", System.Data.DbType.String, this.Code);

            using (var recordset = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                while (!recordset.EoF)
                {
                    itemCodes.Add(recordset.FieldValue("ItemCode").ToString());
                    recordset.MoveNext();
                }
            }

            return itemCodes;
        }

        /// <summary>
        /// Gets the active categories.
        /// </summary>
        /// <returns>List of active categories</returns>
        private IList<string> GetActiveCategories()
        {
            IList<string> categoryCodes = new List<string>();
            var sql = new SqlHelper();
            sql.Builder.AppendLine("SELECT T0.Code FROM [@XX_CTPC] T0");
            sql.Builder.AppendLine("WHERE T0.U_Parent = @Code AND  ISNULL(T0.U_Active, 'N') = 'Y'");
            sql.AddParameter("@Code", System.Data.DbType.String, this.Code);

            using (var recordset = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                while (!recordset.EoF)
                {
                    categoryCodes.Add(recordset.FieldValue("Code").ToString());
                    recordset.MoveNext();
                }
            }

            return categoryCodes;
        }

        /// <summary>
        /// Gets the inactive categories.
        /// </summary>
        /// <returns>List of inactive categories</returns>
        private IList<string> GetInactiveCategories()
        {
            IList<string> categoryCodes = new List<string>();
            var sql = new SqlHelper();
            sql.Builder.AppendLine("SELECT T0.Code FROM [@XX_CTPC] T0");
            sql.Builder.AppendLine("WHERE T0.U_Parent = @Code AND  ISNULL(T0.U_Active, 'N') = 'N'");
            sql.AddParameter("@Code", System.Data.DbType.String, this.Code);

            using (var recordset = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                while (!recordset.EoF)
                {
                    categoryCodes.Add(recordset.FieldValue("Code").ToString());
                    recordset.MoveNext();
                }
            }

            return categoryCodes;
        }

        #endregion Methods
    }
}
