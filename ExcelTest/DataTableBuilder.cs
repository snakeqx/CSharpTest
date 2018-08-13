using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using OfficeOpenXml;
using System.IO;

namespace ConsoleExcelTest
{
    class DataTableBuilder
    {
        public DataTable Table = new DataTable();
        public List<string> ColumnNames = new List<string>();
      
        /// <summary>
        /// Constructor by already existing Worksheet
        /// </summary>
        /// <param name="worksheet">The already loaded ExcelWorkSheet</param>
        public DataTableBuilder(ExcelWorksheet worksheet)
        {
            Table = WorksheetToTable(worksheet);
            GetColumnName();
        }

        /// <summary>
        /// Constructor by a Excel file
        /// </summary>
        /// <param name="filePath">The full excel file name</param>
        /// <param name="page">The sheet number, default is 1</param>
        public DataTableBuilder(string filePath, int page=1)
        {
            try
            {
                FileInfo existingFile = new FileInfo(filePath);
                ExcelPackage package = new ExcelPackage(existingFile);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[page];//选定 指定页
                Table = WorksheetToTable(worksheet);
            }
            catch (Exception ex){throw(ex);}
            GetColumnName();
        }

        private void GetColumnName()
        {
            foreach(DataColumn col in Table.Columns)
            {
                ColumnNames.Add(col.ColumnName);
            }
        }

        private static string GetString(object obj)
        {
            try { return obj.ToString(); }
            catch (Exception) {return "ERROR";}
        }

        private DataTable WorksheetToTable(ExcelWorksheet worksheet)
        {
            int maxRows = worksheet.Dimension.End.Row;
            int maxCols = worksheet.Dimension.End.Column;

            DataTable dt = new DataTable(worksheet.Name);
            DataRow dr = null;

            for (int i = 1; i <= maxRows; i++)
            {
                if (i > 1)
                    dr = dt.Rows.Add();
                for (int j = 1; j <= maxCols; j++)
                {
                    if (i == 1) dt.Columns.Add(GetString(worksheet.Cells[i, j].Value));
                    else dr[j - 1] = GetString(worksheet.Cells[i, j].Value);
                }// end of for j
            }// end of for i
            return dt;
        }
    }
}
