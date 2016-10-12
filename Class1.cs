using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace ClassLibrary1
{
    public class Class1
    {
        const string sexcelconnectionstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\kapil.sharma\\Desktop\\OrderTable.xlsx;Extended Properties='Excel 8.0;HDR=No'";
        const string ssqlconnectionstring = "Data Source=CYG155\\SQLEXPRESS;Initial Catalog = Training Orders Engine Sandbox;Integrated Security=SSPI;";
        public void importdatafromexcel()
        {
            string ssqltable = "OrdersCopy";
            string myexceldataquery = "select * from [sheet1$]";
            try
            {
                //create our connection strings       
                //execute a query to erase any previous data from our destination table
                string sclearsql = "delete from " + ssqltable;
                SqlConnection sqlconn = new SqlConnection(ssqlconnectionstring);
                SqlCommand sqlcmd = new SqlCommand(sclearsql, sqlconn);
                sqlconn.Open();
                sqlcmd.ExecuteNonQuery();
                sqlconn.Close();
                Console.WriteLine("Connection Successfull");
                //series of commands to bulk copy data from the excel file into our sql table
                OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring);
                OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                oledbconn.Open();
                OleDbDataReader dr = oledbcmd.ExecuteReader();
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = ssqltable;
                while (dr.Read())
                {
                    bulkcopy.WriteToServer(dr);
                }

                oledbconn.Close();
            }
            catch (Exception e)
            {
                //handle exception
                Console.WriteLine("Exception Handled\n\n" + e);
            }

        }
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
            }
            catch (IndexOutOfRangeException e)
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
        
        public void customers()
        {

        }
        public void orders()
        {
            
            //SqlConnection sqlconn = new SqlConnection(ssqlconnectionstring);
            //SqlCommand sqlcommand_reader = new SqlCommand("select * from [dbo].[OrdersCopy]", sqlconn);
            //sqlconn.Open();
            //SqlDataReader sql_dr = sqlcommand_reader.ExecuteReader();
            //sqlconn.Close();


            string exlconString = sexcelconnectionstring;
            OleDbConnection exlconn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\kapil.sharma\\Desktop\\OrderTable.xlsx;Extended Properties='Excel 8.0;HDR=Yes'"); 
            OleDbCommand exlcommand_reader = new OleDbCommand("select * from [sheet1$]", exlconn);
            exlconn.Open();
            OleDbDataReader exl_dr = exlcommand_reader.ExecuteReader();


            while (exl_dr.Read())
            {
                Console.WriteLine("For exl_dr: " + exl_dr[0].ToString());

                //SQL connection Object here:
                SqlConnection sqlconn = new SqlConnection(ssqlconnectionstring);
                SqlCommand sqlcommand_reader = new SqlCommand("select * from [dbo].[OrdersCopy]", sqlconn);
                sqlconn.Open();
                SqlDataReader sql_dr = sqlcommand_reader.ExecuteReader();
                

                while (sql_dr.Read())
                {
                    if (sql_dr[0].ToString() == exl_dr[0].ToString())
                    {
                        Console.WriteLine("sql_dr: " + sql_dr[0].ToString()+ "\texl_dr: "+exl_dr[0].ToString());
                        Console.WriteLine("Updation");
                        break;

                    }
                    else
                    {
                        Console.WriteLine("sql_dr: " + sql_dr[0].ToString() + "\texl_dr: " + exl_dr[0].ToString());
                        Console.WriteLine("Insertion");
                    }
                }
                sqlconn.Close();
            }
        }
        public void order_status()
        {

        }    
        public static void Main(string[] args)
        {
            Class1 obj = new Class1();
            //Console.WriteLine("Hello World");
            //obj.importdatafromexcel();
            //Console.WriteLine("Program ended");
            //obj.readFromExcel();

            char choice = 'n';
            int option = 2;
            do
            {
                Console.WriteLine("Choose the operation you wanna do:\n1.Customer\n2.Order\n3.Order Status");
                //option = (int)Console.Read();
                if (option == 1)
                {
                    obj.customers();
                }
                else if (option == 2)
                {
                    obj.orders();
                }
                else if (option == 3)
                {
                    obj.order_status();
                }
                else
                {
                    Console.WriteLine("Invalid Input");
                }
                Console.Write("Do you wish to continue: ");
                //choice = (char)Console.Read();

            } while (choice == 'Y' || choice == 'y');

            Console.ReadLine();
        }

    }
}
