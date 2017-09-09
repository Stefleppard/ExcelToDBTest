using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography;

namespace ExcelToDB_AutoTest
{
    class DelFromDb
    {
        public static void DelAllDb(string connection)
        {
            using (SqlConnection conn = new SqlConnection())
            {
                conn.ConnectionString = connection;
                conn.Open();
                SqlCommand deleteAll = new SqlCommand("DELETE FROM InsertDataTable",conn);
                Console.WriteLine("Deleting all rows" + deleteAll.ExecuteNonQuery());
                conn.Close();
            }
        }
    }
}
