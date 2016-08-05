using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using NPOI.SS.UserModel;
using System.Runtime.InteropServices;
using NPOI.XSSF.UserModel;
using GemBox.Spreadsheet;

namespace DemoExcel
{
    static class Program
    {
        private static string dir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
        static void Main()
        {
            string s = ReadLine();
            CheckSelection(s);          
        }

        private static void CheckSelection(string selection)
        {
            string readLine;
            switch (selection)
            {
                case "1":
                    EpPlus();
                    readLine = ReadLine();
                    CheckSelection(readLine);
                    break;
                case "2":
                    InteropExcel();
                    readLine = ReadLine();
                    CheckSelection(readLine);
                    break;
                case "3":
                    NpoiExcel();
                    readLine = ReadLine();
                    CheckSelection(readLine);
                    break;
                case "4":
                    GemBoxSpreadsSheet();
                    readLine = ReadLine();
                    CheckSelection(readLine);
                    break;
                case "5":
                    break;
                default:
                    readLine = ReadLine();
                    CheckSelection(readLine);
                    break;
            }
        }

        private static string ReadLine()
        {
            Console.WriteLine("Seleccione libreria de Excel a testear");
            Console.WriteLine("1: EPPLUS");
            Console.WriteLine("2: Office.Interop.Excel");
            Console.WriteLine("3: NPOI");
            Console.WriteLine("4: GemBox.SpreadsSheet");
            Console.WriteLine("5: Salir");
            return Console.ReadLine();
        }

        // Process Excel with GemBox.Spreadsheet
        private static void GemBoxSpreadsSheet()
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(dir + "/Excel.xlsx");
            foreach (GemBox.Spreadsheet.ExcelWorksheet sheet in ef.Worksheets)
            {
                foreach (GemBox.Spreadsheet.ExcelRow row in sheet.Rows)
                {
                    foreach (ExcelCell cell in row.AllocatedCells)
                    {
                        if (cell.Value != null)
                            Console.WriteLine(cell.Value);
                    }
                }
            }
            stopwatch.Stop();
            Console.WriteLine(String.Format("Time:  {0}", stopwatch.Elapsed));
        }

        // Process Excel with NPOI Dll
        private static void NpoiExcel()
            {
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                XSSFWorkbook hssfwb;
                using (var file = new FileStream(dir + "/Excel.xlsx", FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new XSSFWorkbook(file);
                }
                ISheet sheet = hssfwb.GetSheetAt(0);
                for (int row = 0; row <= sheet.PhysicalNumberOfRows; row++)
                {
                    if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                    {
                        for (int i = 0; i < sheet.LastRowNum; i++)
                        {
                            Console.WriteLine(sheet.GetRow(row).GetCell(i).StringCellValue);
                        }
                    }
                }
                stopwatch.Stop();
                Console.WriteLine(String.Format("Time:  {0}", stopwatch.Elapsed));
            }

        // Process Excel with EPPlus Dll
        private static void EpPlus()
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            FileStream stream = File.Open("Excel.xlsx", FileMode.Open, FileAccess.Read);
            using (var package = new ExcelPackage(stream))
            {
                var currentSheet = package.Workbook.Worksheets;
                var workSheet = currentSheet.First();
                var noOfCol = workSheet.Dimension.End.Column;
                var noOfRow = workSheet.Dimension.End.Row;
                for (int x = 1; x < noOfRow +1 ; x++)
                {
                    for(int y = 1; y < noOfCol +1; y++)
                    {
                        Console.WriteLine(workSheet.Cells[x, y].Value);
                    }
                }
            }
            stream.Flush();
            stream.Close();
            stopwatch.Stop();
            Console.WriteLine(String.Format("Time:  {0}", stopwatch.Elapsed));
        }

        // Process Excel with Interop Microsoft Dll
        private static void InteropExcel()
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(dir + "/Excel.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
            ExcelScanIntenal(wb);
            // Finish Process
            ClearMemory(app);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Scan the selected Excel workbook and store the information in the cells
        /// for this workbook in an object[,] array. Then, call another method
        /// to process the data.
        /// </summary>
        private static void ExcelScanIntenal(Workbook workBookIn)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            //
            // Get sheet Count and store the number of sheets.
            //
            int numSheets = workBookIn.Sheets.Count;

            //
            // Iterate through the sheets. They are indexed starting at 1.
            //
            for (int sheetNum = 1; sheetNum < numSheets + 1; sheetNum++)
            {
                Worksheet sheet = (Worksheet)workBookIn.Sheets[sheetNum];

                //
                // Take the used range of the sheet. Finally, get an object array of all
                // of the cells in the sheet (their values). You can do things with those
                // values. See notes about compatibility.
                //
                Range excelRange = sheet.UsedRange;
                object[,] valueArray = (object[,])excelRange.get_Value(
                    XlRangeValueDataType.xlRangeValueDefault);

                //
                // Do something with the data in the array with a custom method.
                //
                foreach(var obj in valueArray)
                {
                    Console.WriteLine(obj.ToString());
                }
                Marshal.ReleaseComObject(sheet);
            }
            stopwatch.Stop();
            Console.WriteLine(String.Format("Time:  {0}",stopwatch.Elapsed));
        }

        // To Kill process of Interop Dll
        private static void ClearMemory(Application excelApp)
        {
            excelApp.DisplayAlerts = false;
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }
    }
}
