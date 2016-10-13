using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data;

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


            Boolean flag = true;
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
                        flag = true;
                        
                        break;

                    }
                    else
                    {
                        Console.WriteLine("sql_dr: " + sql_dr[0].ToString() + "\texl_dr: " + exl_dr[0].ToString());
                        flag = false;
                        
                    }
                }
                if (flag == true)
                {
                    Console.WriteLine("Updation");

                }
                else
                {
                    Console.WriteLine("Insertion");
                    SqlCommand cmd = new SqlCommand("Insert into [dbo].[OrdersCopy] values (@OrderID, @CustomerID, @OrderStatusID, @OrderDate, @CurrencyCode, @WarehouuseID, @ShipMethodID, @OrderTypeID, @PriceTypeID, @FirstName, @MiddleName, @LastName, @NameSuffix, @Company, @Address1, @Address2, @Address3, @City, @State, @Zip, @Country, @County, @Email, @Phone, @Notes, @Total, @SubTotal, @TaxTotal, @ShippingTotal, @DiscountTotal, @DiscountPercent, @WeightTotal, @CreatedDate, @ModifiedDate, @CreatedBy, @ModifiedBy");
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = sqlconn;
                    for (int i = 0; i < 36; i++)
                    {
                        cmd.Parameters.AddWithValue("@OrderID", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@CustomerID", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@OrderStatusID", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@OrderDate", DateTime.ParseExact(exl_dr[i].ToString(), "dd/MM/yyyy", ));
                        cmd.Parameters.AddWithValue("@CurrencyCode", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@WarehouuseID", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@ShipMethodID", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@OrderTypeID", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@PriceTypeID", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@FirstName", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@MiddleName", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@LastName", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@NameSuffix", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@Company", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@Address1", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@Address2", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@Address3", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@City", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@State", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@Zip", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@Country", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@County", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@Email", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@Phone", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@Notes", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@Total", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@SubTotal", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@TaxTotal", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@ShippingTotal", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@DiscountTotal", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@DiscountPercent", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@WeightTotal", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@CreatedDate", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@ModifiedDate", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@CreatedBy", Int32.Parse(exl_dr[i].ToString()));
                        cmd.Parameters.AddWithValue("@ModifiedBy", Int32.Parse(exl_dr[i].ToString()));
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
