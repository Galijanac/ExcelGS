using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System;

namespace ExcelGS
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"";
            int sheetNumber = 2;
            int readThisColumn = 1;

            var xlApp = new Excel.Application();
            var xlWorkBookRead = xlApp.Workbooks.Open(path, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            var xlWorkSheetRead = (Excel.Worksheet)xlWorkBookRead.Worksheets.get_Item(sheetNumber);
            var rangeRead = xlWorkSheetRead.UsedRange;

            int rowNumber = rangeRead.Rows.Count;
            int columnNumber = rangeRead.Columns.Count;

            Console.WriteLine((rangeRead.Cells[1, readThisColumn] as Excel.Range).Value2);

            string previousSheet = "";
            int currentSheetNumber = xlApp.Sheets.Count;

            for(int i = 2; i < columnNumber; i++)
            {
                string currentSheet = (rangeRead.Cells[i, readThisColumn] as Excel.Range).Value2;

                if (previousSheet.Equals(currentSheet))
                {
                    //write in it
                }
                else
                {
                    //create new sheet and write in it 
                }
            }
        }
    }
}
