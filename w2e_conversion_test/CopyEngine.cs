using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace w2e_conversion_test
{
    public class CopyEngine
    {
        public CopyEngine(string wordFile, string excelFile) 
        {
            try
            {
                //OPEN NEW INSTANCE OF WORD
                Word._Application objWordApp = new Word.Application();
                objWordApp.Visible = false;
                if (objWordApp == null)
                {
                    Console.WriteLine("Word could not be started. Check that your office installation and project references are correct.");
                    return;
                }
                Word._Document objDoc = objWordApp.Documents.Open(wordFile); //opens specified file in Word instance, stores it in objDoc
                if (objDoc.Tables.Count == 0) //quickfail if document has no tables
                {
                    Console.WriteLine("This document contains no tables");
                    objWordApp.Quit(); //closes Word instance
                    return;
                }

                //IF THE EXCEL FILE HAS ALREADY BEEN MADE, DELETES IT
                if (File.Exists(excelFile))
                {
                    File.Delete(excelFile);
                }

                //OPEN NEW INSTANCE OF EXCEL
                Excel._Application objExcelApp = new Excel.Application(); //creates instance of Excel
                objExcelApp.Visible = false; // false = headless true = visible UI
                Excel._Workbook workbook = objExcelApp.Workbooks.Add(1); //adds a single workbook
                Excel._Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1]; //creates a worksheet within the workbook
                if (worksheet == null) //if excel is acting wonky - bails
                {
                    Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
                    return;
                }
                else
                {
                    Console.WriteLine("And here we go!\nCopy process has been started!");
                }
                int excelRow = 1; //counter for excel rows (so the loops don't continuously overwrite the generated excel data)
                foreach (Word.Table table in objDoc.Tables)
                {
                    string tableHeader = table.Cell(1, 1).Range.Text.Trim(); //grabs the table header and trims whitespace
                    if (tableHeader.StartsWith("CEE")) //we're looking for the CEE tables - nothing else.
                    {
                        for (int row = 1; row <= table.Rows.Count; row++)
                        {
                            for (int col = 1; col <= table.Columns.Count; col++)
                            {
                                try
                                {
                                    worksheet.Cells[excelRow, col] = objExcelApp.WorksheetFunction.Clean(table.Cell(row, col).Range.Text);
                                }
                                catch (Exception)
                                {
                                    continue; // if an exception is thrown, 99% of the time it's due to merged cells in the word table.
                                }
                            }
                            excelRow++;
                        }
                    }

                }
                // Save the excel file 
                workbook.SaveAs(excelFile, Excel.XlFileFormat.xlWorkbookDefault);
                objExcelApp.Workbooks.Close();
                objExcelApp.Application.Quit();

                //Close the Word document
                objWordApp.Documents.Close();
                objWordApp.Quit();
                Console.WriteLine("\nWord document table contents copied to excel file: " + excelFile);
                Console.WriteLine("\nPress ANY Key To Start Conversion");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex);
            }
        }
    }
}
