using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Data;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace ConsoleExcelTest
{
    class CapaSsme
    {
        public DataTableBuilder NcmDataTable;
        public ExcelFileBuilder NcmExcelFile;
        public string ConfigFilePath = @"./config.xml";

        public string EntryDate ="Entry date";
        public string DeliveryDate ="Delivery date";
        public string CloseDate ="Close date";
        public string SourceId ="Source ID";
        public string CreateDate ="Create date";
        public string Reporter = "Reporter";
        public string DefectivePartMaterialNumber ="Defective part material number";
        public string DefectivePartDescription = "Defective part description";
        public string DefectivePartSerialNumber ="Defective part serial number";
        public string Supplier = "Supplier";
        public string Suppliername = "Suppliername";
        public string ProblemDescription ="Problem description";
        public string PCategory = "PCategory";
        public string Responsibility = "Responsibility";
        public string Conclusion = "Conclusion";
        public string ActionCodeFactor = "Action code factor";
        public string FixTime="Fix time (decimal)";
        public string Creator = "Creator";
        public string SystemMaterialNumber = "System material number";
        public string SystemDescription="System description";
        public string strSystemDescriptionNormalized = "System description normalized(string)";
        public string SystemShortDescription = "System short description";
        public string SystemSerialNumber = "System serial number";
        public string ComponentMaterialNumber = "Component material number";
        public string ComponentDescription="Component description";
        public string ComponentSerialNumber = "Component serial number";
        public string Source = "Source";
        public string FilwCount="File count";

        public Dictionary<string, string> NormalizedPoductName;
        public List<string> ListofInstinctProductName;


        private bool isConfigFileExists = false;




        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="page"></param>
        public CapaSsme(string filePath, int page = 1)
        {
            NcmDataTable = new DataTableBuilder(filePath, page);
            InitializeDataTable();
            if (File.Exists(ConfigFilePath))
            {
                ReadConfiXml();
                MatchProductName();
                isConfigFileExists = true;
            }
            else CreateConfigXml();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        public CapaSsme(ExcelWorksheet workSheet)
        {
            NcmDataTable = new DataTableBuilder(workSheet);
            InitializeDataTable();
            if (File.Exists(ConfigFilePath))
            {
                ReadConfiXml();
                MatchProductName();
                isConfigFileExists = true;
            }
            else CreateConfigXml();
        }

        private bool InitializeDataTable()
        {
            // Add a new col with normalized product name
            NcmDataTable.Table.Columns.Add(strSystemDescriptionNormalized, typeof(System.String));
            string tempProductNameNormal;
            // the local r is a reference, so changing r will change NCMDataTable.Table.Rows
            foreach(DataRow r in NcmDataTable.Table.Rows)  
            {
                tempProductNameNormal = r[SystemDescription].ToString().ToUpper();
                
                tempProductNameNormal = NormalizeString(tempProductNameNormal);
                r[strSystemDescriptionNormalized] = tempProductNameNormal;
                
            }

            return true;
        }

        private bool MatchProductName(string configFilePath=@"./config.xml")
        {
            // analyze the product name and remove dedunt
            foreach (DataRow r in NcmDataTable.Table.Rows)
            {
                if (ListofInstinctProductName.Contains(r[SystemDescription])) continue;
                ListofInstinctProductName.Add(r[SystemDescription].ToString());
            }

            // judge if the config file exits

            return true;
        }

        private void CreateConfigXml()
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            List<string> instinctProductNames = new List<string>();
            // get instinctProductNames
            foreach (DataRow r in NcmDataTable.Table.Rows)
            {
                if (instinctProductNames.Contains(r[strSystemDescriptionNormalized])) continue;
                instinctProductNames.Add(r[strSystemDescriptionNormalized].ToString());
            }

            foreach (string i in instinctProductNames)
            {
                if (i == "") dict.Add("N.A.", "N.A.");
                else dict.Add(i, i);
            }
            
            XElement el = new XElement("root", dict.Select(kv => new XElement(kv.Key, kv.Value)));
            el.Save(ConfigFilePath);
        }

        private void ReadConfiXml()
        {
            XElement rootElement = XElement.Parse("<root><key>value</key></root>");
            foreach (var el in rootElement.Elements())
            {
                NormalizedPoductName.Add(el.Name.LocalName, el.Value);
            }
        }

        private static string NormalizeString(string str)
        {
            return str.Replace(" ", "").
                // Be carefule \u00a0 is a "NO_BRAKE_SPACE".
                Replace("\u00a0", "").
                Replace("\u0028", "").
                // Chinese SPACE
                Replace("\u3000", "").
                Replace("\u0029", "");
        }

    }
}
