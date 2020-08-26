using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.CSharp;

namespace EtabsPlugin2017
{
    class Excel
    {
        public string FilePath;
        // constructor
        public Excel(string filePath)
        {
            FilePath = filePath;
        }
        // class method GetExcel
        public List<string>[] GetExcel()
        {

            Application excel = new Application();
            excel.Visible = true;

            Workbook workbook = excel.Workbooks.Open(FilePath);
            Worksheet sheet =  excel.Worksheets.get_Item(1);



            List<string>[] excelData = SetTitleAndListValues(sheet);
            excel.Quit();
           return excelData;
        }

        /// <summary>
        /// read data from an existing excel file
        /// </summary>
        /// <param name="sheet"></param> excel sheet
        /// <returns></returns> excel data as an array
        private List<string>[] SetTitleAndListValues(Worksheet sheet)
        {
            int colCount = sheet.UsedRange.Columns.Count;
            int rowCount = sheet.UsedRange.Rows.Count;

             

            // define a jagged array
            List<string>[] data = new List<string>[colCount];
            for (int column = 1; column <= colCount; column ++)
            {
                int row;
                row = sheet.Cells[column][3].End(XlDirection.xlDown).Row;
                if (row > rowCount)
                {
                    data[column] = null;
                    continue;
                }
                List<string> myList = new List<string>();
                for (int irow = 1; irow <= row-1;irow++)
                {
                   
                    myList.Add(sheet.Cells[column][irow+1].value.ToString());
                       
                }
                data[column-1] = myList;
                

            }
            string a = data[1][2];
            return data;
        }
    }
}
