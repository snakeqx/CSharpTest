using System;
using System.Data;
using System.Xml;

namespace ExcelTest
{
    class Program
    {
        static void Main(string[] args)
        {
            XmlDictHelper xml = new XmlDictHelper(@"/Users/qianxin/Downloads/a.xml");
            xml.AddData("SOMATOM go.Fuck", "Fuck");
            xml.SaveXml(@"/Users/qianxin/Downloads/b.xml");

        }
    }
}
