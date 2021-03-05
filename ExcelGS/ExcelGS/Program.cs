using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System;
using System.Runtime.InteropServices;

namespace ExcelGS
{
    class Program
    {
        public static Excel.Application xlApp;
        public static Excel.Workbook xlWorkBookRead;
        public static Excel.Worksheet xlWorkSheetRead;
        public static Excel.Range rangeRead;

        public static int rowNumberRead;
        public static int columnNumberRead;

        public static Excel.Worksheet xlWorkSheetWrite;

        static void Main(string[] args)
        {
            //@"C:\Users\Nikola\Downloads\text1.xlsx"
            Console.WriteLine("Input path");
            string path = Console.ReadLine();
            // 2
            Console.WriteLine("Input sheet number");
            int sheetNumber = int.Parse(Console.ReadLine());
            // 1
            Console.WriteLine("Input row");
            int readThisColumn = int.Parse(Console.ReadLine());

            try
            {
                xlApp = new Excel.Application();
                xlWorkBookRead = xlApp.Workbooks.Open(path, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                xlWorkSheetRead = (Excel.Worksheet)xlWorkBookRead.Worksheets.get_Item(sheetNumber);
                rangeRead = xlWorkSheetRead.UsedRange;

                rowNumberRead = rangeRead.Rows.Count;
                columnNumberRead = rangeRead.Columns.Count;

                Console.WriteLine((rangeRead.Cells[1, readThisColumn] as Excel.Range).Value2);

                string previousSheet = "";
                int currentSheetNumber = xlApp.Sheets.Count;
                int counter = 1;

                for (int i = 2; i < rowNumberRead; i++)
                {
                    string currentSheet = ((rangeRead.Cells[i, readThisColumn] as Excel.Range).Value2).Replace('/', '-');
                    Console.WriteLine(i);
                    if (previousSheet.Equals(currentSheet))
                    {

                        WriteInSheet(i, counter);

                        counter++;
                    }
                    else
                    {
                        counter = 1;
                        xlWorkSheetWrite = (Excel.Worksheet)xlApp.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        xlWorkSheetWrite.Name = currentSheet;
                        previousSheet = currentSheet;

                        WriteInSheet(i, counter);

                        counter++;
                    }
                }
                string p = path.Replace(".xlsx", "") + DateTime.Now.ToString(" dd.MM.yyyy hh.mm.ss") + ".xlsx";
                xlApp.ActiveWorkbook.SaveAs(p);
            }

            finally
            {
                xlWorkBookRead.Close(true, null, null);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheetWrite);
                Marshal.ReleaseComObject(xlWorkSheetWrite);
                Marshal.ReleaseComObject(xlWorkSheetRead);
                Marshal.ReleaseComObject(xlWorkBookRead);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        private static void WriteInSheet(int oldIndex,int newIndex)
        {
            for(int i = 1; i <= columnNumberRead; i++)
            {
                xlWorkSheetWrite.Cells[newIndex, i] = (rangeRead.Cells[oldIndex, i] as Excel.Range).Value2;
            }
        }
    }
}
