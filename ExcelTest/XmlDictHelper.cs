using System;
using System.Collections;
using System.Data;
using System.Xml;
using System.IO;
using System.Collections.Generic;

namespace ExcelTest
{
    public class XmlDictHelper
    {
        public Dictionary<string, string> Names = new Dictionary<string, string>();
        private string FirstLevelName = "root";
        private string SecondLevelName = "item";
        private string ThirdLevelName1 = "key";
        private string ThirdLevelName2 = "value";

        /// <summary>
        /// Read a file to import the xml
        /// </summary>
        /// <param name="filePath">File path.</param>
        public XmlDictHelper(string filePath)
        {
            if (!File.Exists(filePath)) throw new FileNotFoundException();
            ReadFileToFillDictionary(filePath);
        }

        /// <summary>
        /// Create an xml (in memory) by 2 strings to fill in the dictionary.
        /// </summary>
        /// <param name="key">Key.</param>
        /// <param name="value">Value.</param>
        public XmlDictHelper(string key, string value)
        {
            try{Names.Add(key, value);}
            catch(Exception ex){throw ex;}
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="T:ExcelTest.XmlHelper"/> class.
        /// </summary>
        public XmlDictHelper(){}

        /// <summary>
        /// Saves the xml.
        /// </summary>
        /// <returns><c>true</c>, if xml was saved, <c>false</c> otherwise.</returns>
        /// <param name="filePath">File path.</param>
        /// <param name="isForceDelete">If set to <c>true</c> is force delete.</param>
        public bool SaveXml(string filePath, bool isForceDelete=true)
        {
            // check and deal if file already exists
            if (File.Exists(filePath))
            {
                if (!isForceDelete) throw new Exception("File is already existed. To force save use isForceDelete=true");
                else File.Delete(filePath);
            }

            // Initial xml
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode header = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", null);
            xmlDoc.AppendChild(header);
            XmlElement root = xmlDoc.CreateElement(FirstLevelName);
            // export Names to xml and save file
            foreach(KeyValuePair<string, string> kv in Names)
            {
                XmlElement xn = InsertItem(kv, xmlDoc);
                root.AppendChild(xn);
            }
            xmlDoc.AppendChild(root);
            xmlDoc.Save(filePath);
            return true;
        }

        private XmlElement InsertItem(KeyValuePair<string, string> kv, XmlDocument xmlDoc)
        {
            XmlElement xn = xmlDoc.CreateElement(SecondLevelName);
            xn.AppendChild(GetXmlNode(xmlDoc, ThirdLevelName1, kv.Key));
            xn.AppendChild(GetXmlNode(xmlDoc, ThirdLevelName2, kv.Value));
            return xn;
        }

        private XmlNode GetXmlNode(XmlDocument xmlDoc, string name, string value)
        {
            XmlElement xn = xmlDoc.CreateElement(name);
            xn.InnerText = value;
            return xn;
        }

        /// <summary>
        /// Adds the data.
        /// </summary>
        /// <returns><c>true</c>, if data was added, <c>false</c> if data is updated.</returns>
        /// <param name="key">Key.</param>
        /// <param name="value">Value.</param>
        /// <param name="isForceUpdate">If set to <c>true</c> is force update when key is duplicated.</param>
        public bool AddData(string key, string value, bool isForceUpdate=true)
        {
            if (IsKeyInDictionary(key)){
                if (!isForceUpdate) throw new Exception("Key alrady exists! Maybe use isForceUpdate=true to force update data.");
                else
                {
                    Names[key] = value;
                    return false;
                }
            }
            Names.Add(key, value);
            return true;
        }

        /// <summary>
        /// Reads the file to fill dictionary.
        /// </summary>
        /// <returns><c>true</c>, if file to fill dictionary was  read, <c>false</c> otherwise.</returns>
        /// <param name="filePath">File path.</param>
        protected bool ReadFileToFillDictionary(string filePath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.Load(filePath);
                XmlNodeList itemNodeList = xmlDoc.SelectNodes('/' + FirstLevelName + '/' + SecondLevelName);
                if (itemNodeList == null) return false;
                foreach (XmlNode nl in itemNodeList)
                {
                    string key = nl.FirstChild.InnerXml;
                    string value = nl.LastChild.InnerXml;
                    AddData(key, value);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private bool IsKeyInDictionary(string key)
        {
            if (Names.ContainsKey(key)) return true;
            else return false;
        }



    }// end of class
}
