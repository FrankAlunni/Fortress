// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ProgramCatalogObjectAdapter.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   Defines the ProgramCatalogObjectAdapter type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using System.Data;
using System.Text;
using B1C.SAP.DI.BusinessAdapters.Administration;
using B1C.SAP.DI.Helpers;

namespace B1C.SAP.DI.BusinessAdapters.Catalog
{
    using System;
    using System.Collections.Generic;
    using Components;
    using Inventory;
    using SAPbobsCOM;
    using Utility.Helpers;
    using B1C.SAP.DI.Model.EventArguments;

    /// <summary>
    /// An Instance of a Program Catalog Object
    /// </summary>
    public abstract class ProgramCatalogObjectAdapter : SapObjectAdapter
    {
        #region Fields

        /// <summary>
        /// The active items not in pricing.
        /// </summary>
        private IList<string> activeItemCodesNotInPricing;

        /// <summary>
        /// The active items not in pricing.
        /// </summary>
        private IList<ItemAdapter> activeItemsNotInPricing;

        /// <summary>
        /// The inactive items in pricing.
        /// </summary>
        private IList<string> inactiveItemCodesInPricing;

        /// <summary>
        /// The inactive items in pricing.
        /// </summary>
        private IList<ItemAdapter> inactiveItemsInPricing;

        /// <summary>
        /// The active items not in pricing.
        /// </summary>
        private IList<string> activeItemCodes;

        /// <summary>
        /// The active items not in pricing.
        /// </summary>
        private IList<ItemAdapter> activeItems;

        /// <summary>
        /// The inactive items in pricing.
        /// </summary>
        private IList<string> inactiveItemCodes;

        /// <summary>
        /// The inactive items in pricing.
        /// </summary>
        private IList<ItemAdapter> inactiveItems;

        /// <summary>
        /// A list of Active Categories
        /// </summary>
        private IList<string> activeCategoryCodes;

        /// <summary>
        /// A list of Inactive Categories
        /// </summary>
        private IList<string> inactiveCategoryCodes;

        /// <summary>
        /// A list of Categories adapters
        /// </summary>
        private IList<ProgramCategoryAdapter> activeCategories;

        #endregion Fields

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="ProgramCatalogObjectAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        protected ProgramCatalogObjectAdapter(Company company)
            : base(company)
        {
        }
        #endregion Constructors

        #region Delegates

        /// <summary>
        /// The handler for the progress changed
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The ProgressChangedEvent arguments.</param>
        public delegate void ProgressChangedEventHandler(object sender, ProgressChangedEventArgs e);

        /// <summary>
        /// The handler for the statusbar changed
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">The StatusbarChangedEventArgs arguments.</param>
        public delegate void StatusbarChangedEventHandler(object sender, StatusbarChangedEventArgs e);
        #endregion Delegates

        #region Events
        /// <summary>
        /// Occurs when [progress change].
        /// </summary>
        public event ProgressChangedEventHandler ProgressChanged;

        /// <summary>
        /// Occurs when [statusbar change].
        /// </summary>
        public event StatusbarChangedEventHandler StatusbarChanged;
        #endregion Events

        #region Properties

        /// <summary>
        /// Gets the active items not in pricing.
        /// </summary>
        /// <value>The active items not in pricing.</value>
        public virtual IList<string> ActiveItemCodesNotInPricing
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
        /// Gets the active items not in pricing.
        /// </summary>
        /// <value>The active items not in pricing.</value>
        public IList<ItemAdapter> ActiveItemsNotInPricing
        {
            get
            {
                if (this.activeItemsNotInPricing == null)
                {
                    this.activeItemsNotInPricing = new List<ItemAdapter>();

                    foreach (string itemCode in this.ActiveItemCodesNotInPricing)
                    {
                        var item = new ItemAdapter(this.Company, itemCode);
                        this.activeItemsNotInPricing.Add(item);
                    }
                }

                return this.activeItemsNotInPricing;
            }
        }

        /// <summary>
        /// Gets the inactive item codes in pricing.
        /// </summary>
        /// <value>The inactive item codes in pricing.</value>
        public virtual IList<string> InactiveItemCodesInPricing
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
        /// Gets the inactive items in pricing.
        /// </summary>
        /// <value>The inactive items in pricing.</value>
        public IList<ItemAdapter> InactiveItemsInPricing
        {
            get
            {
                if (this.inactiveItemsInPricing == null)
                {
                    this.inactiveItemsInPricing = new List<ItemAdapter>();

                    foreach (string itemCode in this.InactiveItemCodesInPricing)
                    {
                        var item = new ItemAdapter(this.Company, itemCode);
                        this.inactiveItemsInPricing.Add(item);
                    }
                }

                return this.inactiveItemsInPricing;
            }
        }

        /// <summary>
        /// Gets the active items not in pricing.
        /// </summary>
        /// <value>The active items not in pricing.</value>
        public virtual IList<string> ActiveItemCodes
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
        public IList<ItemAdapter> ActiveItems
        {
            get
            {
                if (this.activeItems == null)
                {
                    this.activeItems = new List<ItemAdapter>();

                    foreach (string itemCode in this.ActiveItemCodes)
                    {
                        var item = new ItemAdapter(this.Company, itemCode);
                        this.activeItems.Add(item);
                    }
                }

                return this.activeItems;
            }
        }

        /// <summary>
        /// Gets the inactive item codes in pricing.
        /// </summary>
        /// <value>The inactive item codes in pricing.</value>
        public virtual IList<string> InactiveItemCodes
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
        /// Gets the inactive items in pricing.
        /// </summary>
        /// <value>The inactive items in pricing.</value>
        public IList<ItemAdapter> InactiveItems
        {
            get
            {
                if (this.inactiveItems == null)
                {
                    this.inactiveItems = new List<ItemAdapter>();

                    foreach (string itemCode in this.InactiveItemCodes)
                    {
                        var item = new ItemAdapter(this.Company, itemCode);
                        this.inactiveItems.Add(item);
                    }
                }

                return this.inactiveItems;
            }
        }

        /// <summary>
        /// Gets the active items not in pricing.
        /// </summary>
        /// <value>The active items not in pricing.</value>
        public virtual IList<string> ActiveCategoryCodes
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
        public virtual IList<string> InactiveCategoryCodes
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

        /// <summary>
        /// Gets the active categories.
        /// </summary>
        /// <value>The active categories.</value>
        public IList<ProgramCategoryAdapter> ActiveCategories
        {
            get
            {
                if (this.activeCategories == null)
                {
                    this.activeCategories = new List<ProgramCategoryAdapter>();

                    foreach (string code in this.ActiveCategoryCodes)
                    {
                        var category = new ProgramCategoryAdapter(this.Company, code, this.CardCode);
                        this.activeCategories.Add(category);
                    }
                }

                return this.activeCategories;
            }
        }

        /// <summary>
        /// Gets or sets the card code.
        /// </summary>
        /// <value>The card code.</value>
        public string Code { get; set; }

        /// <summary>
        /// Gets or sets the global catalog.
        /// </summary>
        /// <value>The global catalog.</value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the description.
        /// </summary>
        /// <value>The description.</value>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets the global catalog code.
        /// </summary>
        /// <value>The global catalog code.</value>
        public string GlobalCatalogCode { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="ProgramCatalogAdapter"/> is active.
        /// </summary>
        /// <value><c>true</c> if active; otherwise, <c>false</c>.</value>
        public bool Active { get; set; }

        /// <summary>
        /// Gets or sets the start date.
        /// </summary>
        /// <value>The start date.</value>
        public DateTime StartDate { get; set; }

        /// <summary>
        /// Gets or sets the end date.
        /// </summary>
        /// <value>The end date.</value>
        public DateTime EndDate { get; set; }

        /// <summary>
        /// Gets or sets the card code.
        /// </summary>
        /// <value>The card code.</value>
        public string CardCode { get; set; }
        #endregion Properties

        #region Methods
        /// <summary>
        /// Adds the new active item to pricing.
        /// </summary>
        /// <returns>A string with all the failures</returns>
        public string AddNewItemsToPricing()
        {
            string outMessage = string.Empty;

            int itemsCount = this.ActiveItemCodesNotInPricing.Count;
            if (itemsCount == 0)
            {
                return outMessage;
            }

            int currentItem = 0;
            foreach (string itemCode in this.ActiveItemCodesNotInPricing)
            {
                try
                {
                    currentItem++;

                    var item = new ItemAdapter(this.Company, itemCode);
                    this.OnProgressChanged(new ProgressChangedEventArgs(currentItem, string.Format("Processing record {0} of {1} - {2} ...", currentItem, itemsCount, item.Item.ItemName)));
                    item.UpdateProgramPricing(this.CardCode);
                }
                catch (Exception ex)
                {
                    outMessage += string.Format("[Program {0}, Item {1}] - {2}{3}", this.CardCode, itemCode, ex.Message, Environment.NewLine);
                }
            }

            return outMessage;
        }

        /// <summary>
        /// Recalculates the pricingfor all the active items in a catalog.
        /// </summary>
        /// <returns>A string with all the failures</returns>
        public string RecalculatePricing()
        {
            string outMessage = string.Empty;
            int itemsCount = this.ActiveItemCodes.Count;

            int currentItem = 0;
            foreach (string itemCode in this.ActiveItemCodes)
            {
                try
                {
                    currentItem++;
                    var item = new ItemAdapter(this.Company, itemCode);
                    this.OnProgressChanged(new ProgressChangedEventArgs(currentItem, string.Format("Processing record {0} of {1} - {2} ...", currentItem, itemsCount, item.Item.ItemName)));
                    item.UpdateProgramPricing(this.CardCode);
                }
                catch (Exception ex)
                {
                    outMessage += string.Format("[Program {0}, Item {1}] - {2}{3}", this.CardCode, itemCode, ex.Message, Environment.NewLine);
                }
            }

            return outMessage;
        }

        /// <summary>
        /// Removes the invalid items.
        /// </summary>
        /// <returns>Remove the non active items</returns>
        public bool DeActivateInvalidItems()
        {
            RecordsetAdapter.Execute(this.Company, string.Format("exec oz_CatalogRemoveInvalidItems '{0}' ", this.CardCode));
             
            return true;
        }

        /// <summary>
        /// Removes the inactive items from pricing.
        /// </summary>
        /// <returns>A string with all the failures</returns>
        public string RemoveInactiveItemsFromPricing()
        {
            string outMessage = string.Empty;

            // Get all the items in the catalog active but not in program pricing
            IList<ItemAdapter> items = this.InactiveItemsInPricing;

            int itemsCount = items.Count;

            int currentItem = 0;
            foreach (ItemAdapter item in items)
            {
                try
                {
                    currentItem++;
                    this.OnProgressChanged(new ProgressChangedEventArgs(currentItem, string.Format("Removing from program pricing record {0} of {1} - {2} ...", currentItem, itemsCount, item.Item.ItemName)));
                    item.RemoveFromPartnerPricing(this.CardCode);
                }
                catch (Exception ex)
                {
                    outMessage += string.Format("[Program {0}, Item {1}] - {2}{3}", this.CardCode, item.ItemCode, ex.Message, Environment.NewLine);
                }
            }

            return outMessage;
        }

        /// <summary>
        /// Raises the ProgressChanged event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.DI.Components.ProgressChangedEventArgs"/> instance containing the event data.</param>
        protected virtual void OnProgressChanged(ProgressChangedEventArgs e)
        {
            this.ProgressChanged(this, e);
        }

        /// <summary>
        /// Raises the StatusbarChanged event.
        /// </summary>
        /// <param name="e">The <see cref="B1C.SAP.DI.Components.StatusbarChangedEventArgs"/> instance containing the event data.</param>
        protected virtual void OnStatusbarChanged(StatusbarChangedEventArgs e)
        {
            this.StatusbarChanged(this, e);
        }

        /// <summary>
        /// Releases the COM Objects.
        /// </summary>
        protected override void Release()
        {
            return;
        }

        /// <summary>
        /// Gets the active items not in pricing.
        /// </summary>
        /// <returns>The list of items not in the pricing but in the catalog</returns>
        private IList<string> GetActiveItemsNotInPricing()
        {
            IList<string> itemCodes = new List<string>();
            var sql = new SqlHelper();
            sql.Builder.AppendLine("SELECT T0.ItemCode");
            sql.Builder.AppendLine("FROM dbo.oz_ItemCatalogActive(@CardCode) T0");
            sql.Builder.AppendLine("LEFT JOIN OSCN T1 ON T1.CardCode = @CardCode AND T0.ItemCode = T1.ItemCode");
            sql.Builder.AppendLine("WHERE T1.ItemCode  IS NULL");
            sql.AddParameter("@CardCode", System.Data.DbType.String, this.CardCode);

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
            sql.Builder.AppendLine("LEFT JOIN dbo.oz_ItemCatalogActive(@CardCode) T1");
            sql.Builder.AppendLine("ON T0.CardCode = @CardCode AND T0.ItemCode = T1.ItemCode");
            sql.Builder.AppendLine("WHERE T1.ItemCode  IS NULL AND T0.CardCode = @CardCode");
            sql.AddParameter("@CardCode", System.Data.DbType.String, this.CardCode);

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
        /// Gets the active items.
        /// </summary>
        /// <returns>
        /// The list of active items
        /// </returns>
        private IList<string> GetActiveItems()
        {
            IList<string> itemCodes = new List<string>();
            var sql = new SqlHelper();
            sql.Builder.AppendLine("SELECT T0.ItemCode");
            sql.Builder.AppendLine("FROM dbo.oz_ItemCatalogActive(@CardCode) T0");
            sql.AddParameter("@CardCode", System.Data.DbType.String, this.CardCode);

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
        /// Gets the inactive items.
        /// </summary>
        /// <returns>
        /// The list of items not active
        /// </returns>
        private IList<string> GetInactiveItems()
        {
            IList<string> itemCodes = new List<string>();
            var sql = new SqlHelper();
            sql.Builder.AppendLine("SELECT T0.ItemCode");
            sql.Builder.AppendLine("FROM dbo.oz_ItemCatalogInactive(@CardCode) T0");
            sql.AddParameter("@CardCode", System.Data.DbType.String, this.CardCode);

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
            sql.Builder.AppendLine("SELECT T0.Code FROM [@XX_CTPC] T0 ");
            sql.Builder.AppendLine("WHERE T1.U_CardCode = @CardCode AND ISNULL(T1.U_Active, 'N') = 'Y'");
            sql.Builder.AppendLine("AND  ISNULL(T0.U_Active, 'N') = 'Y'");
            sql.AddParameter("@CardCode", System.Data.DbType.String, this.CardCode);

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
            sql.Builder.AppendLine("SELECT T0.Code FROM [@XX_CTPC] T0 ");
            sql.Builder.AppendLine("WHERE T1.U_CardCode = @CardCode AND ISNULL(T1.U_Active, 'N') = 'N'");
            sql.Builder.AppendLine("AND  ISNULL(T0.U_Active, 'N') = 'N'");
            sql.AddParameter("@CardCode", System.Data.DbType.String, this.CardCode);

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
        /// Creates the catalog.
        /// </summary>
        /// <param name="catalogName">Name of the catalog.</param>
        /// <returns>True if the catalog is created successfully</returns>
        public bool CreateCatalog(string catalogName)
        {
            string programCode = this.CardCode;

            if (string.IsNullOrEmpty(programCode))
            {
                throw new Exception("Program code is missing.");
            }

            int imageThumbnailSize;
            int imageSmallSize;
            int imageLargeSize;
            int userKey = UserAdapter.Current(this.Company).InternalKey;

            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "ImageThumbnailSize", out imageThumbnailSize);
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "ImageSmallSize", out imageSmallSize);
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "ImageLargeSize", out imageLargeSize);

            // Copy the Catalog
            string query = "SELECT * FROM [@XX_CTCM] WHERE U_Active='Y'";
            using (var catalogRecordset = new RecordsetAdapter(this.Company, query))
            {
                if (!catalogRecordset.EoF)
                {
                    catalogRecordset.MoveFirst();

                    // Add Catalog
                    string code = new TableFieldAdapter(this.Company, "@XX_CTPM", "DocEntry").NextCode(
                        string.Empty, string.Empty);
                    int docEntry = Convert.ToInt32(code);

                    string globalCode = catalogRecordset.FieldValue("Code").ToString();
                    code = new TableFieldAdapter(this.Company, "@XX_CTPM", "Code").NextCode("PCM", "00000");
                    catalogName = string.Format("{0} catalog", catalogName);

                    var stringBuilder = new StringBuilder();
                    stringBuilder.Append("INSERT INTO [@XX_CTPM]([Code],[Name],[DocEntry],");
                    stringBuilder.Append("[Canceled],[Object],[LogInst],[UserSign],");
                    stringBuilder.Append("[Transfered],[CreateDate],[CreateTime],");
                    stringBuilder.Append("[UpdateDate],[UpdateTime],[DataSource],");
                    stringBuilder.Append("[U_CardCode],[U_GlbCatCd],[U_Descript],");
                    stringBuilder.Append("[U_OverrSz],[U_ThumbSz],[U_SmallSz],");
                    stringBuilder.Append("[U_LargeSz],[U_Active],[U_StartDt],[U_EndDt])");
                    stringBuilder.Append("VALUES('{0}','{1}',{2},'N','ProgCatUDO',Null,{3},");
                    stringBuilder.Append("'N','{4}',1000,Null,Null,'I','{5}','{6}','{1}',");
                    stringBuilder.Append("'N',{7},{8},{9},'N','{4}',Null)");

                    query = string.Format(
                        stringBuilder.ToString(),
                        code,
                        catalogName.Replace("'", "''"),
                        docEntry,
                        userKey,
                        DateTime.Today.ToShortDateString(),
                        programCode,
                        globalCode,
                        imageThumbnailSize,
                        imageSmallSize,
                        imageLargeSize);

                    RecordsetAdapter.Execute(this.Company, query);
                    TableAdapter.SetDocEntry(this.Company, (docEntry + 1), "ProgCatUDO");
                    new TableFieldAdapter(this.Company, "@XX_CTCM", "Name").MoveTranslation(globalCode, "@XX_CTPM", code);

                    query = string.Format("SELECT * FROM [@XX_CTCC] WHERE U_CatCode='{0}' AND U_Active='Y'", globalCode);
                    using (var catalogRecordet = new RecordsetAdapter(this.Company, query))
                    {
                        if (!catalogRecordet.EoF)
                        {
                            this.OnStatusbarChanged(new StatusbarChangedEventArgs("The catalog has been set successfully.", false));

                            string newCatalogCode = code;

                            code = new TableFieldAdapter(this.Company, "@XX_CTPC", "DocEntry").NextCode(string.Empty, string.Empty);
                            docEntry = Convert.ToInt32(code);
                            int progressIndicator = 0;

                            while (!catalogRecordet.EoF)
                            {
                                progressIndicator++;

                                code = new TableFieldAdapter(this.Company, "@XX_CTPC", "Code").NextCode("PCC", "00000");
                                string categoryName = catalogRecordet.FieldValue("Name").ToString().Replace("'", "''");
                                string description =
                                    catalogRecordet.FieldValue("U_Descript").ToString().Replace("'", "''");
                                string parent = catalogRecordet.FieldValue("U_Parent").ToString();
                                string active = catalogRecordet.FieldValue("U_Active").ToString();
                                string globalCatCode = catalogRecordet.FieldValue("Code").ToString();

                                this.OnProgressChanged(new ProgressChangedEventArgs(progressIndicator, string.Format("Creating category: {0} ...", categoryName)));

                                // Add the categories 
                                // 0 = Code
                                // 1 = Name
                                // 2 = DocEntry
                                // 3 = User key
                                // 4 = Current Date
                                // 5 = CatCode
                                // 6 = Parent
                                // 7 = Description
                                // 8 = Active
                                // 9 = Global Category Code
                                stringBuilder = new StringBuilder();
                                stringBuilder.Append("INSERT INTO [@XX_CTPC]([Code],[Name],[DocEntry],");
                                stringBuilder.Append("[Canceled],[Object],[LogInst],[UserSign],");
                                stringBuilder.Append("[Transfered],[CreateDate],[CreateTime],[UpdateDate],");
                                stringBuilder.Append("[UpdateTime],[DataSource],[U_CatCode],[U_Parent],");
                                stringBuilder.Append("[U_Descript],[U_Active],[U_StartDt],[U_EndDt], [U_GlbCatCd]) ");
                                stringBuilder.Append("VALUES('{0}','{1}',{2},'N','ProgCatCtUDO',Null,{3},");
                                stringBuilder.Append("'N','{4}',1000,Null,Null,'I','{5}','{6}','{7}','{8}',");
                                stringBuilder.Append("'{4}',null, '{9}')");

                                query = string.Format(
                                    stringBuilder.ToString(),
                                    code,
                                    categoryName.Replace("'", "''"),
                                    docEntry,
                                    userKey,
                                    DateTime.Today.ToShortDateString(),
                                    newCatalogCode,
                                    parent,
                                    description.Replace("'", "''"),
                                    active,
                                    globalCatCode);

                                RecordsetAdapter.Execute(this.Company, query);

                                TableAdapter.SetDocEntry(this.Company, (docEntry + 1), "ProgCatCtUDO");

                                new TableFieldAdapter(this.Company, "@XX_CTCC", "Name").MoveTranslation(globalCatCode, "@XX_CTPC", code);

                                // Add Category Items
                                query =
                                    "SELECT T0.*, T1.[ItemName] FROM [@XX_CTC1] T0 JOIN [OITM] T1 ON T0.[U_ItemCode] = T1.[ItemCode] WHERE T0.[Code] = '{0}' AND T0.[U_Active] = 'Y'";
                                query = string.Format(query, globalCatCode);
                                using (var catalogItemsRecordset = new RecordsetAdapter(this.Company, query))
                                {
                                    if (!catalogItemsRecordset.EoF)
                                    {
                                        // catalogItemsRecordset.MoveFirst();
                                        while (!catalogItemsRecordset.EoF)
                                        {
                                            var item = new ItemAdapter(
                                                this.Company,
                                                catalogItemsRecordset.FieldValue("U_ItemCode").ToString());

                                            // Add the item only if valid for usage
                                            if (string.IsNullOrEmpty(item.IsValidForUsageFromPartner(programCode)))
                                            {
                                                // 0 = Code
                                                // 1 = LineId
                                                // 2 = U_ItemCode
                                                // 3 = U_Active
                                                // 4 = U_StartDt
                                                // 5 = U_EndDt
                                                var helper = new SqlHelper();
                                                helper.Builder.Append("INSERT INTO [@XX_CTP1]([Code],[LineId],[Object],[LogInst],[U_ItemCode],[U_Active],[U_StartDt],[U_EndDt], [U_ItemName], [U_Descript]) ");
                                                helper.Builder.Append("VALUES(@Code,@LineId,'ProgCatCtUDO',Null,@ItemCode,@Active,@StartDate,@EndDate,@ItemName,@Description)");
                                                helper.AddParameter("@Code", DbType.String, code);
                                                helper.AddParameter("@LineId", DbType.Int32, Int32.Parse(catalogItemsRecordset.FieldValue("LineId").ToString()));
                                                helper.AddParameter("@ItemCode", DbType.String, catalogItemsRecordset.FieldValue("U_ItemCode").ToString());
                                                helper.AddParameter("@Active", DbType.String, catalogItemsRecordset.FieldValue("U_Active").ToString());
                                                helper.AddParameter("@StartDate", DbType.DateTime, (DateTime)catalogItemsRecordset.FieldValue("U_StartDt"));
                                                helper.AddParameter("@EndDate", DbType.DateTime, (DateTime)catalogItemsRecordset.FieldValue("U_EndDt"));
                                                helper.AddParameter("@ItemName", DbType.String, catalogItemsRecordset.FieldValue("ItemName").ToString());
                                                helper.AddParameter("@Description", DbType.String, catalogItemsRecordset.FieldValue("ItemName").ToString());

                                                RecordsetAdapter.Execute(this.Company, helper.ToString());

                                                new TableFieldAdapter(this.Company, "@XX_CTC1", "U_ItemName").MoveTranslation(catalogItemsRecordset.FieldValue("Code") + catalogItemsRecordset.FieldValue("LineId").ToString(), "@XX_CTP1", "U_ItemName", code + catalogItemsRecordset.FieldValue("LineId"));
                                                new TableFieldAdapter(this.Company, "@XX_CTC1", "U_Descript").MoveTranslation(catalogItemsRecordset.FieldValue("Code") + catalogItemsRecordset.FieldValue("LineId").ToString(), "@XX_CTP1", "U_Descript", code + catalogItemsRecordset.FieldValue("LineId"));
                                            }

                                            catalogItemsRecordset.MoveNext();
                                        }
                                    }

                                    docEntry++;
                                    catalogRecordet.MoveNext();
                                }
                            }

                            // Update Parent Category Codes
                            query = string.Format("SELECT * FROM [@XX_CTPC] WHERE U_CatCode='{0}'", newCatalogCode);
                            using (var catalogRecordset2 = new RecordsetAdapter(this.Company, query))
                            {
                                if (!catalogRecordset2.EoF)
                                {
                                    // catalogRecordet.MoveFirst();
                                    while (!catalogRecordset2.EoF)
                                    {
                                        query = "UPDATE [@XX_CTPC] SET U_Parent = '{0}' WHERE U_Parent = '{1}'";
                                        query = string.Format(
                                            query,
                                            catalogRecordset2.FieldValue("Code"),
                                            catalogRecordset2.FieldValue("U_GlbCatCd"));
                                        RecordsetAdapter.Execute(this.Company, query);

                                        query = "UPDATE [@XX_CTP1] SET Code= '{0}' WHERE Code = '{1}'";
                                        query = string.Format(
                                            query,
                                            catalogRecordset2.FieldValue("Code"),
                                            catalogRecordset2.FieldValue("U_GlbCatCd"));
                                        RecordsetAdapter.Execute(this.Company, query);

                                        catalogRecordset2.MoveNext();
                                    }
                                }
                            }
                        }

                        this.OnStatusbarChanged(new StatusbarChangedEventArgs("The catalog has been set successfully.", false));
                    }
                }
                else
                {
                    this.OnStatusbarChanged(new StatusbarChangedEventArgs("There aren't any global catalogs active at the moment.", true));
                }
            }

            return true;
        }

        #endregion Methods
    }
}
