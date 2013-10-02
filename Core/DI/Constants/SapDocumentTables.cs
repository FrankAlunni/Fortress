//-----------------------------------------------------------------------
// <copyright file="SapDocumentTables.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
//-----------------------------------------------------------------------
namespace B1C.SAP.DI.Constants
{
    #region Using Directive(s)
    
    using System;
    using System.Collections.Generic;
    
    using SAPbobsCOM;

    #endregion Using Directive(s)
    
    public static class SapDocumentTables
    {
        
        private static IDictionary<BoObjectTypes, string> DocumentNames { get; set; }
        private static IDictionary<BoObjectTypes, string> DocumentTables { get; set; }
        private static IDictionary<BoObjectTypes, int> DocumentCodes { get; set; }
        private static IDictionary<BoObjectTypes, BoAPARDocumentTypes> DocumentAPARDocumentTypes { get; set; }
        private static IDictionary<int, BoObjectTypes> DocumentsByCode { get; set; }
        private static IDictionary<string, BoObjectTypes> DocumentsByTable { get; set; }
        private static IDictionary<BoObjectTypes, string> DocumentReports { get; set; }
        private static IDictionary<BoObjectTypes, int> DocumentForms { get; set; }


        static SapDocumentTables()
        {
            DocumentNames = new Dictionary<BoObjectTypes, string>()
            {
                {BoObjectTypes.oCorrectionPurchaseInvoice,          "A/P Correction Invoice"},
                {BoObjectTypes.oCorrectionPurchaseInvoiceReversal,  "A/P Correction Invoice Reversal"},
                {BoObjectTypes.oPurchaseCreditNotes,                "A/P Credit Memo"},
                {BoObjectTypes.oPurchaseDownPayments,               "A/P Down Payment"},
                {BoObjectTypes.oPurchaseInvoices,                   "A/P Invoice"},
                {BoObjectTypes.oCorrectionInvoice,                  "A/R Correction Invoice"},
                {BoObjectTypes.oCorrectionInvoiceReversal,          "A/R Correction Invoice Reversal"},
                {BoObjectTypes.oCreditNotes,                        "A/R Credit Memo"},
                {BoObjectTypes.oDownPayments,                       "A/R Down Payment"},
                {BoObjectTypes.oInvoices,                           "A/R Invoice"},                
                {BoObjectTypes.oDeliveryNotes,                      "Delivery Notes"},                
                {BoObjectTypes.oPurchaseDeliveryNotes,              "Goods Receipt PO"},
                {BoObjectTypes.oPurchaseReturns,                    "Goods Return"},
                {BoObjectTypes.oPurchaseOrders,                     "Purchase Order"},
                {BoObjectTypes.oReturns,                            "Returns"},
                {BoObjectTypes.oQuotations,                         "Sales Quotation"},
                {BoObjectTypes.oOrders,                             "Sales Order"},
                {BoObjectTypes.oProductionOrders,                   "Production Orders"}
            };
            DocumentTables = new Dictionary<BoObjectTypes, string>()
            {
                {BoObjectTypes.oCorrectionPurchaseInvoice,          "OCPI"},
                {BoObjectTypes.oCorrectionPurchaseInvoiceReversal,  "OCPV"},
                {BoObjectTypes.oPurchaseCreditNotes,                "ORPC"},
                {BoObjectTypes.oPurchaseDownPayments,               "ODPO"},
                {BoObjectTypes.oPurchaseInvoices,                   "OPCH"},
                {BoObjectTypes.oCorrectionInvoice,                  "OCSI"},
                {BoObjectTypes.oCorrectionInvoiceReversal,          "OCSV"},
                {BoObjectTypes.oCreditNotes,                        "ORIN"},
                {BoObjectTypes.oDownPayments,                       "ODPI"},
                {BoObjectTypes.oInvoices,                           "OINV"},                
                {BoObjectTypes.oDeliveryNotes,                      "ODLN"},                
                {BoObjectTypes.oPurchaseDeliveryNotes,              "OPDN"},
                {BoObjectTypes.oPurchaseReturns,                    "ORPD"},
                {BoObjectTypes.oPurchaseOrders,                     "OPOR"},
                {BoObjectTypes.oReturns,                            "ORDN"},
                {BoObjectTypes.oQuotations,                         "OQUT"},
                {BoObjectTypes.oOrders,                             "ORDR"},
                {BoObjectTypes.oProductionOrders,                   "OWOR"}
            };
            DocumentCodes = new Dictionary<BoObjectTypes, int>()
            {
                {BoObjectTypes.oCorrectionPurchaseInvoice,          163},
                {BoObjectTypes.oCorrectionPurchaseInvoiceReversal,  164},
                {BoObjectTypes.oPurchaseCreditNotes,                19},
                {BoObjectTypes.oPurchaseDownPayments,               204},
                {BoObjectTypes.oPurchaseInvoices,                   18},
                {BoObjectTypes.oCorrectionInvoice,                  165},
                {BoObjectTypes.oCorrectionInvoiceReversal,          166},
                {BoObjectTypes.oCreditNotes,                        14},
                {BoObjectTypes.oDownPayments,                       203},
                {BoObjectTypes.oInvoices,                           13},                
                {BoObjectTypes.oDeliveryNotes,                      15},                
                {BoObjectTypes.oPurchaseDeliveryNotes,              20},
                {BoObjectTypes.oPurchaseReturns,                    21},
                {BoObjectTypes.oPurchaseOrders,                     22},
                {BoObjectTypes.oReturns,                            16},
                {BoObjectTypes.oQuotations,                         23},
                {BoObjectTypes.oOrders,                             17},
                {BoObjectTypes.oProductionOrders,                   212}
            };

            DocumentsByCode = new Dictionary<int, BoObjectTypes>()
            {
                {163,   BoObjectTypes.oCorrectionPurchaseInvoice},
                {164,   BoObjectTypes.oCorrectionPurchaseInvoiceReversal},
                {19,    BoObjectTypes.oPurchaseCreditNotes},
                {204,   BoObjectTypes.oPurchaseDownPayments},
                {18,    BoObjectTypes.oPurchaseInvoices},
                {165,   BoObjectTypes.oCorrectionInvoice},
                {166,   BoObjectTypes.oCorrectionInvoiceReversal},
                {14,    BoObjectTypes.oCreditNotes},
                {203,   BoObjectTypes.oDownPayments},
                {13,    BoObjectTypes.oInvoices},
                {15,    BoObjectTypes.oDeliveryNotes},
                {20,    BoObjectTypes.oPurchaseDeliveryNotes},
                {21,    BoObjectTypes.oPurchaseReturns},
                {22,    BoObjectTypes.oPurchaseOrders},
                {16,    BoObjectTypes.oReturns},
                {23,    BoObjectTypes.oQuotations},
                {17,    BoObjectTypes.oOrders},
                {202,   BoObjectTypes.oProductionOrders}

            };

            DocumentAPARDocumentTypes = new Dictionary<BoObjectTypes, BoAPARDocumentTypes>()
            {
                {BoObjectTypes.oCorrectionPurchaseInvoice,          BoAPARDocumentTypes.bodt_CorrectionAPInvoice},
                {BoObjectTypes.oPurchaseCreditNotes,                BoAPARDocumentTypes.bodt_PurchaseCreditNote},
                {BoObjectTypes.oPurchaseInvoices,                   BoAPARDocumentTypes.bodt_PurchaseInvoice},
                {BoObjectTypes.oCorrectionInvoice,                  BoAPARDocumentTypes.bodt_CorrectionARInvoice},
                {BoObjectTypes.oCreditNotes,                        BoAPARDocumentTypes.bodt_CreditNote},
                {BoObjectTypes.oInvoices,                           BoAPARDocumentTypes.bodt_Invoice},                
                {BoObjectTypes.oDeliveryNotes,                      BoAPARDocumentTypes.bodt_DeliveryNote},                
                {BoObjectTypes.oPurchaseDeliveryNotes,              BoAPARDocumentTypes.bodt_PurchaseDeliveryNote},
                {BoObjectTypes.oPurchaseReturns,                    BoAPARDocumentTypes.bodt_PurchaseReturn},
                {BoObjectTypes.oPurchaseOrders,                     BoAPARDocumentTypes.bodt_PurchaseOrder},
                {BoObjectTypes.oReturns,                            BoAPARDocumentTypes.bodt_Return},
                {BoObjectTypes.oQuotations,                         BoAPARDocumentTypes.bodt_Quotation},
                {BoObjectTypes.oOrders,                             BoAPARDocumentTypes.bodt_Order}
            };

            DocumentReports = new Dictionary<BoObjectTypes, string>()
            {
                {BoObjectTypes.oCorrectionPurchaseInvoice,          "APCorrectionInvoice"},
                {BoObjectTypes.oCorrectionPurchaseInvoiceReversal,  "APCorrectionInvoiceReversal"},
                {BoObjectTypes.oPurchaseCreditNotes,                "APCreditMemo"},
                {BoObjectTypes.oPurchaseDownPayments,               "APDownPayment"},
                {BoObjectTypes.oPurchaseInvoices,                   "APInvoice"},
                {BoObjectTypes.oCorrectionInvoice,                  "ARCorrectionInvoice"},
                {BoObjectTypes.oCorrectionInvoiceReversal,          "ARCorrectionInvoiceReversal"},
                {BoObjectTypes.oCreditNotes,                        "ARCreditMemo"},
                {BoObjectTypes.oDownPayments,                       "ARDownPayment"},
                {BoObjectTypes.oInvoices,                           "ARInvoice"},                
                {BoObjectTypes.oDeliveryNotes,                      "DeliveryNote"},                
                {BoObjectTypes.oPurchaseDeliveryNotes,              "GoodsReceiptPO"},
                {BoObjectTypes.oPurchaseReturns,                    "GoodsReturn"},
                {BoObjectTypes.oPurchaseOrders,                     "PurchaseOrder"},
                {BoObjectTypes.oReturns,                            "Return"},
                {BoObjectTypes.oQuotations,                         "SalesQuotation"},
                {BoObjectTypes.oOrders,                             "SalesOrder"},
                {BoObjectTypes.oProductionOrders,                    "ProductionOrder"}
            };

            DocumentForms = new Dictionary<BoObjectTypes, int>()
            {
                {BoObjectTypes.oPurchaseCreditNotes,                181},
                {BoObjectTypes.oPurchaseDownPayments,               65301},
                {BoObjectTypes.oPurchaseInvoices,                   141},
                {BoObjectTypes.oCreditNotes,                        179},
                {BoObjectTypes.oDownPayments,                       65300},
                {BoObjectTypes.oInvoices,                           133},
                {BoObjectTypes.oDeliveryNotes,                      140},
                {BoObjectTypes.oPurchaseDeliveryNotes,              143},
                {BoObjectTypes.oPurchaseReturns,                    182},
                {BoObjectTypes.oPurchaseOrders,                     142},
                {BoObjectTypes.oReturns,                            180},
                {BoObjectTypes.oQuotations,                         149},
                {BoObjectTypes.oOrders,                             139},
                {BoObjectTypes.oProductionOrders,                   65211}
            };

            // Copy the values to the inverted lookups
            //DocumentsByTable = (IDictionary<string, BoObjectTypes>)InvertDictionary(DocumentTables);
            //DocumentsByCode = (IDictionary<int, BoObjectTypes>)InvertDictionary(DocumentCodes);
        }

        /// <summary>
        /// Gets the name of the given document type
        /// </summary>
        /// <param name="objectType">The object type</param>
        /// <returns>The name of the given object type</returns>
        public static string GetDocumentName(BoObjectTypes objectType)
        {
            if (DocumentNames.ContainsKey(objectType))
            {
                return DocumentNames[objectType];
            }

            return string.Empty;
        }

        /// <summary>
        /// Gets the master table of the given object type
        /// </summary>
        /// <param name="objectType">The document object type</param>
        /// <returns>The name of the master table</returns>
        public static string GetDocumentTable(BoObjectTypes objectType)
        {
            if (DocumentTables.ContainsKey(objectType))
            {
                return DocumentTables[objectType];
            }

            return string.Empty;
        }

        /// <summary>
        /// Gets the integer code of the given document object type
        /// </summary>
        /// <param name="objectType">The document object type</param>
        /// <returns>The integer code of the given document object type</returns>
        public static int GetDocumentCode(BoObjectTypes objectType)
        {
            if (DocumentCodes.ContainsKey(objectType))
            {
                return DocumentCodes[objectType];
            }

            return -1;
        }

        /// <summary>
        /// Gets the form code.
        /// </summary>
        /// <param name="objectType">Type of the object.</param>
        /// <returns></returns>
        public static int GetFormCode(BoObjectTypes objectType)
        {
            if (DocumentForms.ContainsKey(objectType))
            {
                return DocumentForms[objectType];
            }

            return -1;
        }

        /// <summary>
        /// Gets the BoAPARDocumentTypes enumerated value corresponding to the given BoObjectTypes type
        /// </summary>
        /// <param name="objectType">The document object type</param>
        /// <returns>The corresponding BoAPARDocumentType</returns>
        public static BoAPARDocumentTypes GetDocumentAPARDocumentType(BoObjectTypes objectType)
        {
            if (DocumentAPARDocumentTypes.ContainsKey(objectType))
            {
                return DocumentAPARDocumentTypes[objectType];
            }

            throw new Exception("Could not find matching BoAPARDocumentTypes value");
        }

        /// <summary>
        /// Gets the Row table for the given object type
        /// </summary>
        /// <param name="objectType">The object type</param>
        /// <returns>The name of the Row table of this document</returns>
        public static string GetDocumentRowTable(BoObjectTypes objectType)
        {
            return GetDocumentTable(objectType).Substring(1, 3) + "1";
        }

        public static string GetReportName(BoObjectTypes objectType)
        {
            return DocumentReports[objectType];
        }


        public static BoObjectTypes GetDocumentTypeFromCode(int documentCode)
        {
            return DocumentsByCode[documentCode]; 
        }

        /*
        public static BoObjectTypes GetDocumentTypeFromTable(string documentTable)
        {
            if (DocumentsByTable.ContainsKey[documentTable])
            {
                return DocumentsByTable[documentTable];
            }

            return null;
        }
        */

        /// <summary>
        /// Inverts a Dictionary object.
        /// NOTE: This must only be used on a dictionary with no duplicate values
        /// </summary>
        /// <param name="dictionary">The input dictionary</param>
        /// <returns>An inverted dictionary</returns>
        private static IDictionary<object, object> InvertDictionary(IDictionary<object, object> dictionary)
        {
            IDictionary<object, object> inverse = new Dictionary<object, object>();

            foreach (object key in dictionary.Keys)
            {
                object value = dictionary[key];

                // Check if the inverted dictionary already contains this value as the key
                if (inverse.ContainsKey(value))
                {
                    throw new Exception("Can not invert a dictionary with duplicate values");
                }
                else
                {
                    inverse[value] = key;
                }
            }

            return inverse;
        }
        
    }
}

