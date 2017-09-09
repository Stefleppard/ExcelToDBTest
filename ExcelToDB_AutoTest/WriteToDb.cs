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
    class WriteToDb
    {
        public static void WriteLnToDb(string connection, List<int> values)
        {
            int date = values[0];
            int price = values[1];
            int value  = values[2];
            using (SqlConnection conn = new SqlConnection())
            {
                conn.ConnectionString = connection;
                conn.Open();
                SqlCommand command = new SqlCommand("Insert into InsertDataTable (Date, Price, Value) Values (@0,@1,@2) ", conn);
                command.Parameters.Add(new SqlParameter("0", date));
                command.Parameters.Add(new SqlParameter("1", Convert.ToDouble(price)));
                command.Parameters.Add(new SqlParameter("2", Convert.ToDouble(value)));
                Console.WriteLine("Insert executed! Total rows affected are " + command.ExecuteNonQuery());
                conn.Close();
            }
        }
    }
}
