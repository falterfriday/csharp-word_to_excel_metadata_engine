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

        public void WriteToExcel(List<Dictionary<string, string>> conversionList, Excel._Worksheet worksheet)
        {
            try
            {
                int excelRow = 2;
                int excelCol = 1;

                foreach (var dictionary in conversionList)
                {
                    foreach (var keyValue in dictionary)
                    {
                        worksheet.Cells[excelRow, excelCol].Formula = keyValue.Value;
                        excelCol++;
                    }
                    excelRow++;
                    excelCol = 1;
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Uh Oh... Something broke in the ExcelWriter.");
            }
            
        }
    }
}
