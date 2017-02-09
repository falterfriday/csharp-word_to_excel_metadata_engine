using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace w2e_conversion_test
{
    public class ConversionEngine
    {
        public ConversionEngine(string copiedExcelFile)
        {
            try
            {
                Console.WriteLine("Conversion process has been started... Hang tight!");
                
                //OPEN THE COPIED EXCEL FILE
                Excel._Application copiedExcelApp = new Excel.Application();
                Excel.Workbook copiedExcelWorkbook = copiedExcelApp.Workbooks.Open(copiedExcelFile);
                Excel.Worksheet copiedExcelWorksheet = copiedExcelWorkbook.Sheets["Sheet1"] as Excel.Worksheet;
                
                int presentRows = copiedExcelWorksheet.UsedRange.Rows.Count; //grabs the number of used rows in the excel worksheet
                
                Excel.Range range = null;

                Parser parser = new Parser();//fires up a Parser object

                for (int row = 1; row <= presentRows; row++)
                {
                    for (int col = 1; col <= 4; col++)
                    {
                        range = copiedExcelWorksheet.Cells[row, col];
                        string cellText = (range.Text.ToString()).Trim();

                        if (cellText != null || cellText != "")
                        {
                            parser.CheckText(cellText, col);
                        }
                    }
                }

                //CLOSE THE COPIED EXCEL FILE (HAVING A TON OF EXCEL INSTANCES RUNNING IN THE BKGRD = NO BUENO)
                copiedExcelApp.Workbooks.Close();
                copiedExcelApp.Application.Quit();
            }
            catch (Exception)
            {
                Console.WriteLine("Uh Oh... Something broke in the Conversion Engine.");
            }
        }
    }
}
