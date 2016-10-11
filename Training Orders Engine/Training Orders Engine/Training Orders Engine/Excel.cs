using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Training_Orders_Engine
{
    class Excel
    {
        public static void excel()
        {
          /*  string Exfilepath = @"C:\Users\siddharth.bhatnagar\Desktop\Customer.xls";
            Excel.Application xlApp = new Excel.Application();
            
            // Open the Excel file.
            // You have pass the full path of the file.
            // In this case file is stored in the Bin/Debug application directory.
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Exfilepath);

            // Get the first worksheet.
            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets.get_Item(1);

            // Get the range of cells which has data.
            Excel.Range xlRange = xlWorksheet.UsedRange;

            // Get an object array of all of the cells in the worksheet with their values.
            object[,] valueArray = (object[,])xlRange.get_Value(
                        Excel.XlRangeValueDataType.xlRangeValueDefault);

            // iterate through each cell and display the contents.
            for (int row = 1; row <= xlWorksheet.UsedRange.Rows.Count; ++row)
            {
                for (int col = 1; col <= xlWorksheet.UsedRange.Columns.Count; ++col)
                {
                    // Print value of the cell to Console.
                    Console.WriteLine(valueArray[row, col].ToString());
                }
            }

            // Close the Workbook.
            xlWorkbook.Close(false);

            // Relase COM Object by decrementing the reference count.
            Marshal.ReleaseComObject(xlWorkbook);

            // Close Excel application.
            xlApp.Quit();

            // Release COM object.
            Marshal.FinalReleaseComObject(xlApp);

            Console.ReadLine();*/
        }
    }
}
