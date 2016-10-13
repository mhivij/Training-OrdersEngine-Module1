using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace ClassLibrary1
{
    class Class2
    {
        public void readFromExcel()
        {
            //Creating connection strings
            string conString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\kapil.sharma\\Desktop\\OrderTable.xlsx;" + "Extended Properties='Excel 8.0;HDR=Yes;'";
            OleDbConnection conn = new OleDbConnection(conString);  
            int col_count = 0;
            OleDbCommand command_reader = new OleDbCommand("select * from [sheet1$]", conn);
            conn.Open();
            OleDbDataReader dr_count = command_reader.ExecuteReader();
            int i = 0;
            //Counting Column from the Excel Sheet
            try
            {
                while (dr_count.Read())
                {
                    while (true)
                    {
                        var rowcol = dr_count[i];
                        i++;
                    }
                }
            }catch(IndexOutOfRangeException e)
            {
                Console.WriteLine("Count = " + i);
                col_count = i;       
            }
            dr_count.Close();
            //Displaying table data on the console
            OleDbDataReader dr = command_reader.ExecuteReader();
            while (dr.Read())
            {
                for (i = 0; i < col_count; i++)
                {
                    var rowcol = dr[i];
                    Console.Write(rowcol + "\t");                 
                }
                Console.Write("\n");
            }
            conn.Close();
        }
        //public static void Main(string[] args)
        //{
        //    Class2 obj = new Class2();
        //    obj.readFromExcel();
        //    Console.ReadLine();
        //}
    
    }
}
