using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Training_Orders_Engine
{
    class Connection
    {
        public static SqlConnection Conn;
        public static void connect()
        {
            Conn = new SqlConnection(Properties.Settings.Default.ConnStr);
            Conn.Open();
            Console.WriteLine("Connection successfull");
        }
    }
}
