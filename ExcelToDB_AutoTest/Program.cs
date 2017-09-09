using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDB_AutoTest
{
    class Program
    {
        static void Main(string[] args)
        {
            DelFromDb.DelAllDb(ConfigurationManager.ConnectionStrings["DBConnection"].ToString());
            ReadExcel.GetExcelFile();
        }
    }
}
