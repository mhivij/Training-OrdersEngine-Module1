using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Training_Orders_Engine
{
    class ConnectAndViewExcel
    {
        public static void excel()
        {
            string Exfilepath = @"C:\Users\siddharth.bhatnagar\Desktop\Customer1.xls";
            DataTable dt = new DataTable();
            //DataRow dr = null;
            OleDbConnection Conn = null;
            OleDbDataAdapter MyCommand = null;
            Conn = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='" + Exfilepath + "';Extended Properties=Excel 8.0;");

            //If you MS Excel 2007 then use below lin instead of above line
            //Conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Exfilepath + "';Extended Properties=Excel 12.0;");

            MyCommand = new OleDbDataAdapter("select * from [Sheet1$]", Conn);
            Conn.Open();
            DataSet DtSet = new DataSet();
            MyCommand.Fill(DtSet, "[Sheet1$]");
            dt = DtSet.Tables[0];
            Console.WriteLine(dt);
        }
    }
}
