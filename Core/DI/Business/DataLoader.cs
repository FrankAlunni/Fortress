//-----------------------------------------------------------------------
// <copyright file="DataLoader.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.DI.Business
{
    using System;
    using System.IO;
    using System.Xml;
    using System.Xml.Serialization;

    using SAPbobsCOM;
    using System.Text;

    /// <summary>
    /// Instance of the Field class.
    /// </summary>
    /// <author>Frank Alunni</author>
    [SerializableAttribute]
    [XmlTypeAttribute(AnonymousType = true)]
    [XmlRootAttribute(Namespace = "", IsNullable = false)]
    public class Field
    {
        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>The field name.</value>
        /// <remarks/>
        [XmlAttributeAttribute]
        public string Name
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        /// <value>The field type.</value>
        /// <remarks/>
        [XmlAttributeAttribute]
        public string Type
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        /// <value>The field value.</value>
        /// <remarks/>
        [XmlAttributeAttribute]
        public string Value
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <returns>Return the object value in its original data type</returns>
        public object GetValue()
        {
            try
            {
                return Convert.ChangeType(this.Value, System.Type.GetType(this.Type));
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// Instance of the Data Loader class.
    /// </summary>
    /// <author>Frank Alunni</author>
    [SerializableAttribute]
    [XmlTypeAttribute(AnonymousType = true)]
    [XmlRootAttribute(IsNullable = false)]
    public class DataLoader
    {
        /// <summary>
        /// Gets or sets the record array.
        /// </summary>
        /// <value>The record array.</value>
        [XmlElementAttribute("Record", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public Record[] Records
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the name of the UDO.
        /// </summary>
        /// <value>The name of the UDO.</value>
        [XmlAttributeAttribute]
        public string UDOName
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the type of the UDO.
        /// </summary>
        /// <value>The type of the UDO.</value>
        /// <remarks/>
        [XmlAttributeAttribute]
        public string UDOType
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the error message.
        /// </summary>
        /// <value>The error message.</value>
        public string ErrorMessage
        {
            get;
            set;
        }

        /// <summary>
        /// Instances from XML String
        /// </summary>
        /// <param name="xml">The xml data</param>
        /// <returns>The dataloader created</returns>
        public static DataLoader InstanceFromXMLString(string xml)
        {
            var encoding = new System.Text.UTF8Encoding();
            var serializer = new XmlSerializer(typeof(DataLoader));

            using (var stream = new MemoryStream(encoding.GetBytes(xml)))
            {               
                return (DataLoader)serializer.Deserialize(stream);             
            }
        }

        /// <summary>
        /// Instances from XML.
        /// </summary>
        /// <param name="filename">The XML filename.</param>
        /// <returns>The dataloader created</returns>
        public static DataLoader InstanceFromXML(string filename)
        {
            // Create an instance of the XmlSerializer specifying type and namespace.
            var serializer = new XmlSerializer(typeof(DataLoader));

            // A FileStream is needed to read the XML document.
            using (var fs = new FileStream(filename, FileMode.Open))
            {
                using (XmlReader reader = new XmlTextReader(fs))
                {
                    return (DataLoader)serializer.Deserialize(reader);
                }
            }
        }

        /// <summary>
        /// Processes the specified company.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <returns>True if the dataloader is processed succesfully</returns>
        public bool Process(Company company)
        {
            // Validate Company object
            if (company == null)
            {
                throw new ArgumentNullException("company", "The company object is not initialized.");
            }

            if (!company.Connected)
            {
                throw new ArgumentException("The company object is not connected.", "company");
            }

            if (company.InTransaction)
            {
                company.EndTransaction(BoWfTransOpt.wf_RollBack);
            }

            CompanyService companyService = company.GetCompanyService();
            GeneralService generalService = null;
            GeneralData parentGeneralData = null;
            GeneralDataCollection childTableGeneralDataCollection = null;
            GeneralData childTableGeneralData = null;

            company.StartTransaction();

            try
            {
                // Retrieve the relevant service
                generalService = companyService.GetGeneralService(this.UDOName);

                foreach (Record headerRecord in this.Records)
                {
                    // Point to the Header of the MD UDO
                    parentGeneralData = (GeneralData)generalService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);

                    foreach (Field headerField in headerRecord.Fields)
                    {
                        // Insert values to the Header properties
                        object fieldValue = headerField.GetValue();
                        if (fieldValue != null)
                        {
                            parentGeneralData.SetProperty(headerField.Name, fieldValue);
                        }
                    }

                    if (headerRecord.ChildTables != null)
                    {
                        foreach (ChildTable childTable in headerRecord.ChildTables)
                        {
                            // Insert Values to the Lines tables
                            childTableGeneralDataCollection = parentGeneralData.Child(childTable.Name);
                            if (childTable.Rows != null)
                            {
                                foreach (Row childRow in childTable.Rows)
                                {
                                    childTableGeneralData = childTableGeneralDataCollection.Add();
                                    foreach (Field childField in childRow.Fields)
                                    {
                                        object fieldValue = childField.GetValue();
                                        if (fieldValue != null)
                                        {
                                            childTableGeneralData.SetProperty(childField.Name, fieldValue);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Add - MD UDO Header and Line Data to DB
                    generalService.Add(parentGeneralData);
                }

                if (company.InTransaction)
                {
                    company.EndTransaction(BoWfTransOpt.wf_Commit);
                }

                return true;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                this.ErrorMessage = string.Format("COM Exception: {0}", ex.Message);
                if (company.InTransaction)
                {
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }

                return false;
            }
            catch (Exception ex)
            {
                this.ErrorMessage = ex.Message;
                if (company.InTransaction)
                {
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }

                return false;
            }
            finally
            {
                if (childTableGeneralData != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(childTableGeneralData);
                }

                if (childTableGeneralDataCollection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(childTableGeneralDataCollection);
                }

                if (parentGeneralData != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(parentGeneralData);
                }

                if (generalService != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(generalService);
                }

                if (companyService != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(companyService);
                }

                childTableGeneralData = null;
                childTableGeneralDataCollection = null;
                parentGeneralData = null;
                generalService = null;
                companyService = null;

                GC.Collect();
            }
        }
    }

    /// <summary>
    /// Instance of the Record class.
    /// </summary>
    /// <author>Frank Alunni</author>
    [SerializableAttribute]
    [XmlTypeAttribute(AnonymousType = true)]
    public class Record
    {
        /// <summary>
        /// Gets or sets the field array.
        /// </summary>
        /// <value>The field array.</value>
        /// <remarks/>
        [XmlElementAttribute("Field")]
        public Field[] Fields
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the child table array.
        /// </summary>
        /// <value>The child table array.</value>
        /// <remarks/>
        [XmlElementAttribute("ChildTables", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public ChildTable[] ChildTables
        {
            get;
            set;
        }
    }

    /// <summary>
    /// Instance of the Child Table class.
    /// </summary>
    /// <author>Frank Alunni</author>
    [SerializableAttribute]
    [XmlTypeAttribute(AnonymousType = true)]
    public class ChildTable
    {
        /// <summary>
        /// Gets or sets the row class array.
        /// </summary>
        /// <value>The row class array.</value>
        /// <remarks/>
        [XmlElementAttribute("Row", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public Row[] Rows
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the name of the child table.
        /// </summary>
        /// <value>The name of the child table.</value>
        /// <remarks/>
        [XmlAttributeAttribute]
        public string Name
        {
            get;
            set;
        }
    }

    /// <summary>
    /// Instance of the Rows class.
    /// </summary>
    /// <author>Frank Alunni</author>
    [SerializableAttribute]
    [XmlTypeAttribute(AnonymousType = true)]
    public class Row
    {
        /// <summary>
        /// Gets or sets the field array.
        /// </summary>
        /// <value>The field array.</value>
        /// <remarks/>
        [XmlElementAttribute("Field")]
        public Field[] Fields
        {
            get;
            set;
        }
    }

    /// <summary>
    /// Instance of the DataLoaderSet class.
    /// </summary>
    /// <author>Frank Alunni</author>
    [SerializableAttribute]
    [XmlTypeAttribute(AnonymousType = true)]
    [XmlRootAttribute(IsNullable = false)]
    public class DataLoaderSet
    {
        /// <summary>
        /// Gets or sets the items.
        /// </summary>
        /// <value>The set items.</value>
        /// <remarks/>
        [XmlElementAttribute("Field", typeof(Field))]
        [XmlElementAttribute("DataLoader", typeof(DataLoader))]
        public object[] Items
        {
            get;
            set;
        }
    }
}