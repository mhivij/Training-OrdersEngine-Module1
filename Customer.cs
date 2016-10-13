using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Training_Orders_Engine
{
    public class Customer
    {
        public static void Main(string[] args)
        {
            Connection.connect();
            new ConnectAndViewExcel().excel();


            //Console.WriteLine("Connection successfull");
            Console.ReadLine();
        }
    }
}
