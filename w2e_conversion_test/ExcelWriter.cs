using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace w2e_conversion_test
{
    public class ExcelWriter
    {
        public ExcelWriter()
        {
            Excel._Application objExcelApp = new Excel.Application();
            objExcelApp.Visible = false;
            Excel._Workbook workbook = objExcelApp.Workbooks.Add(1);
            Excel._Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            if (worksheet == null)
            {
                Console.WriteLine("Worksheet could not be created. Check the ExcelWriter");
                return;
            }
            else
            {
                Console.WriteLine("Writing metadata to new file");
            }
        }

        public void WriteToExcel(List<Dictionary<string, string>> conversionList, Excel._Worksheet worksheet)
        {
            int excelRow = 2;
            int excelCol = 1;
            foreach (var dictionary in conversionList)
            {
                foreach (var keyValue in dictionary)
                {
                    Console.WriteLine(string.Format("{0}: {1}", keyValue.Key, keyValue.Value));
                }
            }
        }
    }
}
