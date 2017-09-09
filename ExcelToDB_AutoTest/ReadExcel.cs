using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       

namespace ExcelToDB_AutoTest
{
    public class ReadExcel
    {
        public static void GetExcelFile()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ConfigurationManager.AppSettings["ExcelFile"]);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            List<int> values = new List<int>();

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    var value = Convert.ToInt32(xlRange.Cells[i, j].Value2);
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        Console.Write(value + "\t");
                        values.Insert(j - 1, value);
                    
                    if (j == 3)
                    {
                        Console.Write("\r\n");
                        WriteToDb.WriteLnToDb(ConfigurationManager.ConnectionStrings["DBConnection"].ToString(),values);
                    }
                }
            }
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

    }
}
