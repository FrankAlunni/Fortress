// --------------------------------------------------------------------------------------------------------------------
// <copyright file="XmlHelper.cs" company="B1C Canada Inc.">
//   Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <summary>
//   XmlHelper Class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace B1C.Utility.Helpers
{
    using System;
    using System.IO;
    using System.Text;
    using System.Xml;
    using System.Xml.Serialization;
    using System.Xml.Xsl;

    /// <summary>
    ///  XmlHelper Class
    /// </summary>
    public class XmlHelper
    {
        /// <summary>
        /// Serializes to string.
        /// </summary>
        /// <param name="workingObject">The working object.</param>
        /// <returns>The serialized string</returns>
        public static string SerializeToString(object workingObject)
        {
            var memoryStream = new MemoryStream();
            var xs = new XmlSerializer(workingObject.GetType());
            var xmlTextWriter = new XmlTextWriter(memoryStream, new UTF8Encoding(false));

            xs.Serialize(xmlTextWriter, workingObject);

            return UTF8ByteArrayToString(memoryStream.ToArray());
        }

        /// <summary>
        /// Deserializes from string.
        /// </summary>
        /// <param name="workingString">The working string.</param>
        /// <param name="workingType">Type of the working.</param>
        /// <returns>The deserialized string.</returns>
        public static object DeserializeFromString(string workingString, Type workingType)
        {
            var memoryStream = new MemoryStream(StringToUTF8ByteArray(workingString));

            // Call Webservice Method (Using a direct call inside the assembly)
            var xs = new XmlSerializer(workingType);
            return xs.Deserialize(memoryStream);
        }

        /// <summary>
        /// Applies the stylesheet.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="xsltFilePath">The XSLT file path.</param>
        /// <param name="outputFilePath">The output file path.</param>
        public static void ApplyStylesheet(XmlDocument document, string xsltFilePath, string outputFilePath)
        {
            // Convert it into General Settings XML
            var xct = new XslCompiledTransform();
            xct.Load(new XmlTextReader(new StreamReader(xsltFilePath)));

            var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, Indent = true };

            XmlWriter writer = XmlWriter.Create(outputFilePath, settings);
            if (writer != null)
            {
                xct.Transform(document, writer);
                writer.Close();
            }
        }

        /// <summary>
        /// UTF8 byte array to string.
        /// </summary>
        /// <param name="characters">The characters.</param>
        /// <returns>The UTF8 String</returns>
        protected static string UTF8ByteArrayToString(byte[] characters)
        {
            var encoding = new UTF8Encoding();
            string constructedString = encoding.GetString(characters);
            return constructedString;
        }

        /// <summary>
        /// Strings to UTF8 byte array.
        /// </summary>
        /// <param name="xmlString">The XML string.</param>
        /// <returns>The UTF8 Byte array</returns>
        protected static byte[] StringToUTF8ByteArray(string xmlString)
        {
            var encoding = new UTF8Encoding();
            byte[] byteArray = encoding.GetBytes(xmlString);
            return byteArray;
        }

        /// <summary>
        /// Validates the XML doc.
        /// </summary>
        /// <param name="xmlFileLocation">The XML file location.</param>
        /// <param name="xsdFileLocation">The XSD file location.</param>
        /// <param name="results">The results of the XSD Validation.</param>
        /// <returns>true if successful, false otherwise</returns>
        public bool ValidateXMLDoc(string xmlFileLocation, string xsdFileLocation, out StringBuilder results)
        {
            bool isValid = true;
            results = new StringBuilder(string.Format("Processing file: {0} with {1}", xmlFileLocation, xsdFileLocation));
            results.AppendLine("Start");

            var lastLine = string.Empty;

            try
            {
                var settings = new XmlReaderSettings();
                settings.Schemas.Add(null, xsdFileLocation);
                settings.ValidationType = ValidationType.Schema;
                var document = new XmlDocument();
                document.Load(xmlFileLocation);
            
                var reader = XmlReader.Create(new StringReader(document.OuterXml), settings);

                while (reader.Read())
                {
                    lastLine = string.Empty;

                    for (int i = 0; i < reader.Depth; i++)
                        lastLine += " ";

                    lastLine += reader.Name + " : " + reader.Value;
                    results.Append(lastLine);
                    results.AppendLine(" ... OK");
                }
            }
            catch (Exception ex)
            {
                results.AppendFormat("Exception encountered on line: {0}\n", lastLine);
                results.AppendLine(ex.Message);
                isValid = false;
            }
            string res = results.ToString();

            return isValid;
        }
    }
}
