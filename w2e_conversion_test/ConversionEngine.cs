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

                // NEW LIST OF DICTIONARIES
                List<Dictionary<string, string>> conversionList = new List<Dictionary<string, string>>();

                
                Excel._Application objExcelApp = new Excel.Application();
                objExcelApp.Visible = true;
                Excel._Workbook workbook = objExcelApp.Workbooks.Add(1);
                Excel._Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
                if (worksheet == null)
                {
                    Console.WriteLine("Worksheet could not be created. Check the ExcelWriter");
                    return;
                }
                
                //NEW PARSER, PASSING LIST
                Parser parser = new Parser();
                string cellText;

                for (int row = 1; row <= presentRows; row++)
                {
                    for (int col = 1; col <= 4; col++)
                    {
                        range = copiedExcelWorksheet.Cells[row, col];
                        cellText = (range.Text.ToString()).Trim();

                        if (cellText != null && cellText != "")
                        {
                            parser.CheckText(conversionList, cellText, col);
                        }
                    }
                }

                ExcelWriter writer = new ExcelWriter();
                writer.WriteToExcel(conversionList, worksheet);
                
                //CLOSE THE COPIED EXCEL FILE (HAVING A TON OF EXCEL INSTANCES RUNNING IN THE BKGRD = NO BUENO)
                copiedExcelApp.Workbooks.Close();
                copiedExcelApp.Application.Quit();
                objExcelApp.Workbooks.Close();
                objExcelApp.Application.Quit();
            }
            catch (Exception)
            {
                Console.WriteLine("Uh Oh... Something broke in the Conversion Engine.");
            }
        }
    }
}
