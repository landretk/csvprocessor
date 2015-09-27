using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace CsvProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;


            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            if (officeType == null)
            {
                //no Excel installed
                Console.WriteLine("Excel not installed, hit any key to quit...");
                Console.ReadLine();
                return;
            }



            xlApp = new Excel.Application();
            
            string currentDirName = Directory.GetCurrentDirectory();
            string[] files = Directory.GetFiles(currentDirName, "*.csv");

            Console.WriteLine("Processing CSV files into XLSX...");

            foreach (string s in files)
            {
                // Create the FileInfo object only when needed to ensure 
                // the information is as current as possible.
                System.IO.FileInfo fi = null;
                try
                {
                    fi = new FileInfo(s);
                }
                catch (FileNotFoundException e)
                {
                    // To inform the user and continue is 
                    // sufficient for this demonstration. 
                    // Your application may require different behavior.
                    Console.WriteLine(e.Message);
                    continue;
                }
                Console.WriteLine("Converting: {0}", s);

                xlWorkBook = xlApp.Workbooks.Open(s,0,true,1,"","",true,Excel.XlPlatform.xlWindows,'\t',true);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                range = xlWorkSheet.get_Range("A7", "A60024");
                try
                {
                    xlApp.DisplayAlerts = false;
                    range.TextToColumns(xlWorkSheet.get_Range("A7"), Excel.XlTextParsingType.xlDelimited, Excel.XlTextQualifier.xlTextQualifierDoubleQuote,
                        false, true, false, false, false, false, false);
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    Console.WriteLine(e.Message);
                    continue;
                }
                    int extension = Path.GetExtension(s).Length;
                    xlWorkBook.SaveAs(currentDirName + "\\" + xlWorkBook.Name.Substring(0, xlWorkBook.Name.Length - extension) + ".xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook);

                    xlWorkBook.Close(false);
                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
            }
            xlApp.Quit();
            releaseObject(xlApp);
#if DEBUG
            Console.WriteLine("\nProcessing complete, hit any key to quit");
            Console.ReadLine();
#else
            Console.WriteLine("\nProcessing complete");
#endif
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        } 
    }
}
