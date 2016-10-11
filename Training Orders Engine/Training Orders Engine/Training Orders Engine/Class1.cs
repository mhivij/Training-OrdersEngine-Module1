using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Training_Orders_Engine
{
    public class Class1
    {
        public static void Main(string[] args)
        {
            Connection.connect();
            ConnectAndViewExcel.excel();
            //Console.WriteLine("Connection successfull");
            Console.ReadLine();
        }
    }
}
