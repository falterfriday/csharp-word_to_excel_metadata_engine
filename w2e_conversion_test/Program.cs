using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;


namespace w2e_conversion_test
{
    class Program
    {
        static void Main(string[] args)
        {
            string wordFile = @"C:\temp\TestDoc.docx";
            string copiedExcelFile = @"C:\temp\TestDoc.xlsx";
            string excelMetaDataFile = @"C:\temp\MetaDataTestDoc.xlsx";
            new CopyEngine(wordFile, copiedExcelFile);
            new ConversionEngine(copiedExcelFile);
            Console.WriteLine("ah... finished!");
            Console.WriteLine("Press ANY Key To Exit");
            Console.ReadLine();
        }
    }
}
