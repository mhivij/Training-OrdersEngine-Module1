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
        public static void connect()
        {
            /* SqlConnection Conn = new SqlConnection(
                                        "Trusted_Connection=yes;" +
                                        "Data Source=172.25.122.106,1433" +
                                        "Initial Catalog=Training Orders Engine Sandbox;"+
                                        "password=P@ssw0rd;"+"user id = sa;");*/
            SqlConnection Conn = new SqlConnection(Properties.Settings.Default.ConnStr);
            //Conn.Open();
            Console.WriteLine("Connection successfull");
        }
    }
}
