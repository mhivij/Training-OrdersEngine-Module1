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
            Console.WriteLine("Customers");
            Console.WriteLine("Enter 1 for customers");
            int value = Console.Read();
            switch (value)
            {
                case 49:
                    new ConnectAndViewExcel().excel();
                    break;
            }
            Console.ReadLine();

        }
        
    }
}
