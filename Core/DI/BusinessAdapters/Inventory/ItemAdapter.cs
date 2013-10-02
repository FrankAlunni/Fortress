//-----------------------------------------------------------------------
// <copyright file="ItemAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------

namespace B1C.SAP.DI.BusinessAdapters.Inventory
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Text;
    using Business;
    using BusinessPartners;
    using Components;
    using Helpers;
    using SAPbobsCOM;
    using Utility.Helpers;
    using Utility.Logging;
    using B1C.SAP.DI.Model.Exceptions;

    /// <summary>
    /// The Item adapter instance
    /// </summary>
    public class ItemAdapter : SapObjectAdapter
    {
        #region Fields
        /// <summary>
        /// Whether or not the item exists in the DB
        /// </summary>
        private readonly bool exists;

        /// <summary>
        /// The Item Code
        /// </summary>
        private readonly string itemCode;

        /// <summary>
        /// The Kitted items quanities
        /// </summary>
        private IDictionary<string, int> kittedItemsQuantities;

        /// <summary>
        /// The kitted items
        /// </summary>
        private IList<string> kittedItems;

        /// <summary>
        /// The Default Vendor Adapter
        /// </summary>
        private BusinessPartnerAdapter defaultVendor;

        /// <summary>
        /// The base item
        /// </summary>
        private Items item;
        #endregion Fields

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="ItemAdapter"/> class.
        /// </summary>
        /// <param name="company">
        /// The SAP company connection
        /// </param>
        /// <param name="code">
        /// The item code.
        /// </param>
        public ItemAdapter(Company company, string code)
            : base(company)
        {
            this.itemCode = code.Trim(' ');
            this.Item = (Items)this.Company.GetBusinessObject(BoObjectTypes.oItems);
            this.exists = this.Item.GetByKey(this.itemCode);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ItemAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="item">The item object.</param>
        public ItemAdapter(Company company, Items item)
            : base(company)
        {
            this.itemCode = item.ItemCode;            
            this.Item = item;            
            this.exists = true;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ItemAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="cardCode">The card code.</param>
        /// <param name="businessPartnerCatalogNumber">The business partner catalog number.</param>
        public ItemAdapter(Company company, string cardCode, string businessPartnerCatalogNumber) : base(company)
        {
            this.itemCode = this.GetItemCodeFromPartnerCode(cardCode, businessPartnerCatalogNumber).Trim(' ');
            this.Item = (Items)this.Company.GetBusinessObject(BoObjectTypes.oItems);
            this.exists = this.Item.GetByKey(this.itemCode);
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets or sets the item.
        /// </summary>
        /// <value>The SAP item.</value>
        public Items Item
        {
            get { return this.item; }
            set { this.item = value; }
        }

        /// <summary>
        /// Gets the ITems' code
        /// </summary>
        public string ItemCode
        {
            get
            {
                return this.itemCode;
            }
        }

        /// <summary>
        /// Gets the available stock of this item (across all warehouses)
        /// </summary>
        /// <value>The available stock.</value>
        public int AvailableStock
        {
            get
            {
                string query = string.Format(
                    @"SELECT [dbo].oz_ItemStockAvailable('{0}')", 
                    this.itemCode);

                using (var rows = new RecordsetAdapter(this.Company, query))
                {
                    if (rows.EoF)
                    {
                        return -1;
                    }

                    rows.MoveFirst();
                    return (int) rows.FieldValue(0);
                }
            }
        }

        #region Protected Properties
        #endregion Protected Properties

        #endregion

        #region Methods

        /// <summary>
        /// Determines whether item is managed by serial numbers.
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if item is managed by serial numbers; otherwise, <c>false</c>.
        /// </returns>
        public bool IsItemSerialized()
        {
            return this.Item.ManageSerialNumbers == BoYesNoEnum.tYES;
        }

        /// <summary>
        /// Checks if the item exists
        /// </summary>
        /// <returns>true if the item exists, false otherwise</returns>
        public bool Exists()
        {
            return this.exists;
        }

        /// <summary>
        /// Defaults the vendor.
        /// </summary>
        /// <returns>The Default Business Partner</returns>
        public virtual BusinessPartnerAdapter DefaultVendor()
        {
            this.Task = string.Format("Getting default vendor for item {0}", this.ItemCode);
            if (this.defaultVendor == null || this.defaultVendor.Document == null)
            {
                string cardCode = this.Item.Mainsupplier;

                if (!string.IsNullOrEmpty(cardCode))
                {
                    this.defaultVendor = new BusinessPartnerAdapter(this.Company, cardCode);
                }
            }

            return this.defaultVendor;
        }

        /// <summary>
        /// Return true of the item is excusive to the card code
        /// </summary>
        /// <param name="cardCode">
        /// The card code.
        /// </param>
        /// <returns>
        /// True of the item is excusive to the card code
        /// </returns>
        public virtual bool IsPartnerExclusive(string cardCode)
        {
            this.Task = string.Format("Checking if item {0} is partner exclusive", this.ItemCode);

            bool isExclusive;

            var builder = new StringBuilder();
            builder.AppendLine(@"SELECT 1 FROM [@XX_EXCLUSIVEITEM]");
            builder.AppendLine(@"WHERE  GetDate() BETWEEN ISNULL(U_From, CAST('19000101' as dateTime)) ");
            builder.AppendFormat(@"AND ISNULL(U_To, CAST('29991231' as dateTime)) AND ISNULL(U_InActive, '') <> 'Y' AND U_ItemCode = '{0}' AND U_CardCode = '{1}'", this.itemCode, cardCode);

            using (var rows = new RecordsetAdapter(this.Company, builder.ToString()))
            {
                isExclusive = !rows.EoF;
            }

            return isExclusive;
        }

        /// <summary>
        /// Determines whether this instance is a BOM.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance is a BOM; otherwise, <c>false</c>.
        /// </returns>
        public bool IsBOM()
        {
            bool isBom = false;
            string query = string.Format("SELECT ItemCode FROM OITM WHERE ItemCode = '{0}' AND TreeType = 'P'", this.ItemCode);

            using (var adapter = new RecordsetAdapter(this.Company, query))
            {
                if (!adapter.EoF)
                {
                    isBom = true;
                }
            }

            return isBom;
        }

        /// <summary>
        /// Return true of the item is an excusive item
        /// </summary>
        /// <returns>
        /// True of the item is an excusive item
        /// </returns>
        public virtual bool IsExclusive()
        {
            this.Task = string.Format("Checking if item {0} is exclusive", this.ItemCode);

            bool isExclusive;

            var builder = new StringBuilder();
            builder.AppendLine(@"SELECT 1 FROM [@XX_EXCLUSIVEITEM]");
            builder.AppendLine(@"WHERE  GetDate() BETWEEN ISNULL(U_From, CAST('19000101' as dateTime)) ");
            builder.AppendFormat(@"AND ISNULL(U_To, CAST('29991231' as dateTime)) AND U_ItemCode = '{0}' AND ISNULL(U_Inactive, '') <> 'Y'", this.itemCode);

            using (var rows = new RecordsetAdapter(this.Company, builder.ToString()))
            {
                isExclusive = !rows.EoF;
            }

            return isExclusive;
        }

        /// <summary>
        /// Return true of the item is on hold
        /// </summary>
        /// <returns>
        /// True of the item is on hold
        /// </returns>
        public virtual bool IsOnHold()
        {
            this.Task = string.Format("Checking if item {0} is on hold", this.ItemCode);

            bool isOnHold;

            ExchangeHelper exchange = ExchangeHelper.Instance(this.Company);
            if (exchange == null)
            {
                throw new NullReferenceException();
            }

            var builder = new StringBuilder();
            builder.AppendLine("exec sp_executesql ");
            builder.AppendLine("N'SELECT frozenFor FROM OITM WHERE ItemCode = (@P1)',");
            builder.AppendLine("N'@P1 nvarchar(60)', ");
            builder.AppendFormat("N'{0}'", this.itemCode);

            using (var itemRecordset = new RecordsetAdapter(this.Company, builder.ToString()))
            {
                isOnHold = itemRecordset.EoF ||
                    (itemRecordset.FieldValue("frozenFor").ToString() == "Y" ? true : false);
            }

            return isOnHold;
        }

        /// <summary>
        /// Determines whether [is discontinued until inventory exhausted].
        /// </summary>
        /// <returns>
        /// <c>true</c> if [is discontinued until inventory exhausted]; otherwise, <c>false</c>.
        /// </returns>
        public virtual bool IsDiscontinuedUntilInventoryExhausted()
        {
            this.Task = string.Format("Checking if item {0} is discontinued until inventory is exhausted", this.ItemCode);

            bool isDiscontinued = false;

            string query = string.Format(@"SELECT [dbo].oz_ItemStatus('{0}')", this.itemCode);

            using (var rows = new RecordsetAdapter(this.Company, query))
            {
                switch (rows.FieldValue(0).ToString())
                {
                    case "DiscExaustInventory":
                        isDiscontinued = this.AvailableStock <= 0;
                        break;
                    default:
                        break;
                }
            }

            return isDiscontinued;
        }

        /// <summary>
        /// Tests if the item is Expired.
        /// </summary>
        /// <returns>True if the item is expired, false otherwise</returns>
        public virtual bool IsExpired()
        {
            this.Task = string.Format("Checking if item {0} is expired", this.ItemCode);

            var sql =
                new SqlHelper(
                    "SELECT 1 FROM OITM WHERE GetDate() BETWEEN ISNULL(validFrom, CAST('19000101' as dateTime)) AND ISNULL(validTo, CAST('29991231' as datetime)) and ItemCode = @ItemCode");

            sql.AddParameter("@ItemCode", System.Data.DbType.String, this.itemCode);

            using (var rows = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                return rows.EoF;
            }
        }

        /// <summary>Updates or create a business partner catalog number</summary>
        /// <param name="partnerCode">The partner code.</param>
        /// <param name="oldSubstitute"> The old substitute.</param>
        /// <param name="newSubstitute"> The new substitute.</param>
        /// <returns>True if the method runs successfully</returns>
        public bool UpdateBusinessPartnerCatalog(string partnerCode, string oldSubstitute, string newSubstitute)
        {
            this.Task = string.Format("Updating item {0} BP Catalog number for {1} from {2} to {3}", this.ItemCode, partnerCode, oldSubstitute, newSubstitute);

            // Add the new Catalog Item
            var catalog = (AlternateCatNum)this.Company.GetBusinessObject(BoObjectTypes.oAlternateCatNum);

            try
            {
                if (!catalog.GetByKey(this.itemCode, partnerCode, oldSubstitute))
                {
                    catalog.ItemCode = this.itemCode;
                    catalog.Substitute = newSubstitute;
                    catalog.DisplayBPCatalogNumber = BoYesNoEnum.tYES;
                    catalog.CardCode = partnerCode;

                    if (catalog.Add() != 0)
                    {
                        throw new SapException(this.Company);
                    }
                }
                else
                {
                    catalog.Substitute = newSubstitute;

                    if (catalog.Update() != 0)
                    {
                        throw new SapException(this.Company);
                    }
                }
            }
            finally
            {
                COMHelper.Release(ref catalog);
            }

            return true;
        }

        /// <summary>Updates or create a business partner catalog number</summary>
        /// <param name="partnerCode">The partner code.</param>
        /// <returns>True if the method runs successfully</returns>
        public bool UpdateBusinessPartnerCatalog(string partnerCode)
        {
            this.Task = string.Format("Updating item {0} business partner catalog for {1}", this.ItemCode, partnerCode);

            // Add the new Catalog Item
            string query = string.Format("SELECT * FROM OSCN WHERE CardCode = '{0}' AND ItemCode = '{1}'", partnerCode, this.itemCode);
            using (var recordset = new RecordsetAdapter(this.Company, query))
            {
                if (recordset.EoF)
                {
                    var catalog = (AlternateCatNum)this.Company.GetBusinessObject(BoObjectTypes.oAlternateCatNum);
                    try
                    {
                        if (!catalog.GetByKey(this.itemCode, partnerCode, this.itemCode))
                        {
                            catalog.ItemCode = this.itemCode;
                            catalog.Substitute = this.itemCode;
                            catalog.DisplayBPCatalogNumber = BoYesNoEnum.tYES;
                            catalog.CardCode = partnerCode;

                            if (catalog.Add() != 0)
                            {
                                throw new SapException(this.Company);
                            }
                        }
                    }
                    finally
                    {
                        COMHelper.Release(ref catalog);
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Quantities the on stock.
        /// </summary>
        /// <param name="wareHouseCode">The ware house code.</param>
        /// <returns>
        /// The On hand quantity at a specific warehouse
        /// </returns>
        public int QuantityOnStock(string wareHouseCode)
        {
            this.Task = string.Format("Checking quantity on stock for item {0} in warehouse {1} ", this.ItemCode, wareHouseCode);

            SqlHelper sql;

            // Check it the item is a BOM 
            if (this.Item.TreeType == BoItemTreeTypes.iAssemblyTree)
            {
                ThreadedAppLog.WriteLine(string.Format("\t\tItem {0} is a BOM.", this.Item.ItemCode));

                // Check the minimum qty available for the BOM components
                sql = new SqlHelper();
                sql.Builder.AppendLine("SELECT TOP 1 CAST(ISNULL(SUM(T0.OnHand), 0)/ ISNULL(AVG(T1.Quantity), 0) AS INT)  as OnHand");
                sql.Builder.AppendLine("FROM OITW T0 JOIN ITT1 T1 ON T0.ItemCode = T1.Code");
                sql.Builder.AppendLine("WHERE T1.Father = @ItemCode AND WhsCode = @WhsCode");
                sql.Builder.AppendLine("GROUP BY T1.Code ORDER BY CAST(ISNULL(SUM(T0.OnHand), 0)/ ISNULL(AVG(T1.Quantity), 0) AS INT)");
                sql.AddParameter("@ItemCode", System.Data.DbType.String, this.Item.ItemCode);
                sql.AddParameter("@WhsCode", System.Data.DbType.String, wareHouseCode);
            }
            else
            {
                sql = new SqlHelper("SELECT ISNULL(SUM(OnHand), 0) as OnHand FROM OITW WHERE ItemCode = @ItemCode AND WhsCode = @WhsCode");
                sql.AddParameter("@ItemCode", System.Data.DbType.String, this.Item.ItemCode);
                sql.AddParameter("@WhsCode", System.Data.DbType.String, wareHouseCode);
            }

            using (var rs = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                if (!rs.EoF)
                {
                    return Convert.ToInt32(rs.FieldValue("OnHand").ToString());
                }

                ThreadedAppLog.WriteLine("Unable to retrieve the OnHand quantity of item {0}\n{1}", this.ItemCode, sql.ToString());
            }

            return 0;            
        }

        /// <summary>
        /// Prices the specified price list.
        /// </summary>
        /// <param name="priceList">The price list.</param>
        /// <returns>The price of a specific price list</returns>
        public virtual double Price(int priceList)
        {
            this.Task = string.Format("Checking item {0} price in price list {1} ", this.ItemCode, priceList);

            var sql = new SqlHelper();
            sql.Builder.AppendLine("SELECT DISTINCT T0.ItemCode,");
            sql.Builder.AppendLine("CASE ISNULL(SUM((T2.Quantity/T1.Qauntity) * T3.Price), 0)");
            sql.Builder.AppendLine("	WHEN 0 THEN T4.Price");
            sql.Builder.AppendLine("	ELSE ISNULL(SUM((T2.Quantity/T1.Qauntity) * T3.Price), 0)");
            sql.Builder.AppendLine("END AS Price");
            sql.Builder.AppendLine("FROM OITM T0");
            sql.Builder.AppendLine("LEFT JOIN OITT T1 ON T0.ItemCode = T1.Code");
            sql.Builder.AppendLine("LEFT JOIN ITT1 T2 ON T2.Father = T1.Code ");
            sql.Builder.AppendFormat("LEFT JOIN ITM1 T3 ON T2.Code = T3.ItemCode AND T3.PriceList = {0} ", priceList);
            sql.Builder.AppendFormat("JOIN ITM1 T4 ON T0.ItemCode = T4.ItemCode AND T4.PriceList = {0}", priceList);
            sql.Builder.AppendLine("GROUP BY T0.ItemCode, T4.Price ");
            sql.Builder.AppendFormat("HAVING T0.ItemCode = '{0}'", this.ItemCode);

            using(var rs = new RecordsetAdapter(this.Company, sql.ToString()))
            {
                return rs.EoF ? 0.0d : Convert.ToDouble(rs.FieldValue("Price").ToString());
            }
        }

        /// <summary>
        /// Gets the kitted items.
        /// </summary>
        /// <returns>List of kited items</returns>
        public IDictionary<string, int> GetKittedItemQuantities()
        {
            this.Task = string.Format("Getting item {0} kitted quantities", this.ItemCode);

            if (this.kittedItemsQuantities == null)
            {
                this.kittedItemsQuantities = new Dictionary<string, int>();

                using(var rs = new RecordsetAdapter(this.Company, string.Format("SELECT Code, Quantity from [ITT1] where father = '{0}'", this.ItemCode)))
                {
                    while (!rs.EoF)
                    {
                        string code = rs.FieldValue("Code").ToString();
                        int quantity = 1;
                        Int32.TryParse(rs.FieldValue("Quantity").ToString(), out quantity);

                        if (code != null)
                        {
                            this.kittedItemsQuantities.Add(code, quantity);
                        }

                        rs.MoveNext();
                    }
                }
            }

            return this.kittedItemsQuantities;
        }

        /// <summary>
        /// Gets the kitted items.
        /// </summary>
        /// <returns>List of kited items</returns>
        public IList<string> GetKittedItems()
        {
            this.Task = string.Format("Getting item {0} kitted items", this.ItemCode);

            if (this.kittedItems == null)
            {
                this.kittedItems = new List<string>();

                string query = string.Format("SELECT Code from [ITT1] where father = '{0}'", this.ItemCode);

                using (var rs = new RecordsetAdapter(this.Company, string.Format("SELECT Code, Quantity from [ITT1] where father = '{0}'", this.ItemCode)))
                {
                    while (!rs.EoF)
                    {
                        string code = rs.FieldValue("Code").ToString();
                        if (code != null)
                        {
                            this.kittedItems.Add(code);
                        }

                        rs.MoveNext();
                    }
                }
            }

            return this.kittedItems;
        }

        /// <summary>
        /// Partners the catalog number.
        /// </summary>
        /// <param name="partnerCode">The partner code.</param>
        /// <returns>Returns the partners the catalog number.</returns>
        public virtual string PartnerCatalogNumber(string partnerCode)
        {
            this.Task = string.Format("Getting item {0} BP Catalog number for {1}", this.ItemCode, partnerCode);

            string query = string.Format("SELECT SubStitute FROM OSCN WHERE ItemCode = '{0}' AND CardCode = '{1}'", this.itemCode, partnerCode);
            using (var catalogRecordset = new RecordsetAdapter(this.Company, query))
            {
                return !catalogRecordset.EoF ? catalogRecordset.FieldValue("SubStitute").ToString().Trim(' ') : string.Empty;
            }
        }

        /// <summary>
        /// Gets the item code from BP code.
        /// </summary>
        /// <param name="cardCode">The card code.</param>
        /// <param name="businessPartnerCatalogNumber">The business partner catalog number.</param>
        /// <returns>The item code</returns>
        public string GetItemCodeFromPartnerCode(string cardCode, string businessPartnerCatalogNumber)
        {
            this.Task = string.Format("Getting item code from partner code {0} for partner {1}", businessPartnerCatalogNumber, cardCode);

            using (var recordset = new RecordsetAdapter(this.Company, string.Format("SELECT ItemCode from OSCN where Substitute = '{0}' and CardCode = '{1}'", businessPartnerCatalogNumber, cardCode)))
            {
                return !recordset.EoF ? recordset.FieldValue("ItemCode").ToString() : string.Empty;
            }
        }

        /// <summary>
        /// Releases the Documents COM Object.
        /// </summary>
        protected override void Release()
        {
            if (this.defaultVendor != null)
            {
                this.defaultVendor.Dispose();
                this.defaultVendor = null;
            }

            COMHelper.Release(ref this.item);
        }

        /// <summary>
        /// Uploads an image to the given destinations. This method will create a copy of the original, a thumbnail, small image and large image
        /// from the input source image.
        /// </summary>
        /// <param name="sourceFileName">The location of the source image on the filesystem</param>
        /// <param name="destinationFileName">The location (and filename) to place the copy of the original</param>
        /// <param name="thumbnailFileName">The location (and filename) to place the thumbnail</param>
        /// <param name="smallFileName">The location (and filename) to place the small image</param>
        /// <param name="largeFileName">The location (and filename) to place teh large image</param>
        /// <returns>True if the image is uploaded successfully</returns>
        protected bool DoUpload(string sourceFileName, string destinationFileName, string thumbnailFileName, string smallFileName, string largeFileName)
        {
            this.Task = string.Format("Uploading image for item {0}", this.ItemCode);

            int imageThumbnailSize;
            int imageSmallSize;
            int imageLargeSize;
            float imageResolution;

            // Load the sizing values from configuration
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "ImageThumbnailSize", out imageThumbnailSize);
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "ImageSmallSize", out imageSmallSize);
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "ImageLargeSize", out imageLargeSize);
            ConfigurationHelper.GlobalConfiguration.Load(this.Company, "ImageResolution", out imageResolution);

            Image originalImage = Image.FromFile(sourceFileName);

            originalImage.Save(destinationFileName, System.Drawing.Imaging.ImageFormat.Jpeg);

            Image thumbnailImage = Utility.Imaging.Engine.Resize(
                    originalImage, imageThumbnailSize, imageThumbnailSize, imageResolution, Color.White);
            thumbnailImage.Save(thumbnailFileName, System.Drawing.Imaging.ImageFormat.Jpeg);
            thumbnailImage.Dispose();

            originalImage = Image.FromFile(sourceFileName);
            Image smallImage = Utility.Imaging.Engine.Resize(
                    originalImage, imageSmallSize, imageSmallSize, imageResolution, Color.White);
            smallImage.Save(smallFileName, System.Drawing.Imaging.ImageFormat.Jpeg);
            smallImage.Dispose();

            originalImage = Image.FromFile(sourceFileName);
            Image largeImage = Utility.Imaging.Engine.Resize(
                    originalImage, imageLargeSize, imageLargeSize, imageResolution, Color.White);
            largeImage.Save(largeFileName, System.Drawing.Imaging.ImageFormat.Jpeg);
            largeImage.Dispose();

            return true;
        }
        #endregion Methods
    }
}
