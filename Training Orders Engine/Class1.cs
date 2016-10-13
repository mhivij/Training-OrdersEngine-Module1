﻿using System;
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
        DataTable DtExcelData;
        DataTable DtSqlData;
        OleDbConnection ExcelConn;
        OleDbDataAdapter DaExcelcmd;
        SqlConnection Conn;

        public void importdatafromexcel()
        {
            Conn = new SqlConnection(Training_Orders_Engine.Properties.Settings.Default.ConnStr);
            Conn.Open();
            Console.WriteLine("Connection successfull");

        DtExcelData = new DataTable();
            DtSqlData = new DataTable();

            string Exfilepath = @"C:\Users\siddharth.bhatnagar\Desktop\Customer.xlsx";
           //If you MS Excel 2007 then use below lin instead of above line
           ExcelConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Exfilepath + "';Extended Properties='Excel 12.0;hdr=yes;'");

            DaExcelcmd = new OleDbDataAdapter("select * from [Sheet1$]", ExcelConn);
            ExcelConn.Open();
            DaExcelcmd.Fill(DtExcelData);

            SqlCommand com = new SqlCommand("select * from Customer", Conn);
            com.CommandType = CommandType.Text;
            SqlDataAdapter DaSqlcmd = new SqlDataAdapter(com);
            DaSqlcmd.Fill(DtSqlData);

            /*string ssqltable = "OrdersCopy";
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
            }*/

        }
        public void readFromExcel()
        {
            //Creating connection strings
          /*  string conString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\kapil.sharma\\Desktop\\OrderTable.xlsx;" + "Extended Properties='Excel 8.0;HDR=Yes;'";
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
            conn.Close();*/
        }
        
        public void customers()
        {
            int sqlrow = DtSqlData.Rows.Count;
            bool Flag = false;
            int excelrow = DtExcelData.Rows.Count;

            for (int i = 0; i < excelrow; i++)
            {
                for (int j = 0; j < sqlrow; j++)
                {
                    if (DtExcelData.Rows[i][0].ToString() == DtSqlData.Rows[j][0].ToString())
                    {
                        Flag = true;
                        break;
                    }
                    else
                    {
                        Flag = false;
                    }

                }
                if (Flag)
                {
                    SqlCommand cmd = new SqlCommand("UPDATE Customer SET FirstName=@FirstName,MiddleName=@MiddleName, LastName=@LastName, Company=@Company, CustomerTypeID=@CustomerTypeID, CustomerStatusID=@CustomerStatusID, Email=@Email, Phone=@Phone, MainAddress1=@MainAddress1, MainAddress2=@MainAddress2, MainAddress3=@MainAddress3, MainCity=@MainCity, MainState=@MainState, MainZip=@MainZip, MainCountry=@MainCountry, MailAddress1=@MailAddress1, MailAddress2=@MailAddress2, MailAddress3=@MailAddress3, MailCity=@MailCity, MailState=@MailState, MailZip=@MailZip, MailCountry=@MailCountry,CanLogin=@CanLogin, LoginName=@LoginName, BirthDate=@BirthDate, CurrencyCode=@CurrencyCode, LanguageID=@LanguageID, Gender=@Gender, TaxCode=@TaxCode, TaxCodeTypeID=@TaxCodeTypeID, IsSalesTaxExempt=@IsSalesTaxExempt, SalesTaxCode=@SalesTaxCode, IsEmailSubscribed=@IsEmailSubscribed, Notes=@Notes, CreatedDate=@CreatedDate, ModifiedDate=@ModifiedDate, CreatedBy=@CreatedBy, ModifiedBy=@ModifiedBy WHERE CustomerID=@CustomerID",Conn);
                    cmd.Parameters.Add("@CustomerID", SqlDbType.Int).Value = DtExcelData.Rows[i][0];
                    cmd.Parameters.Add("@FirstName", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][1].ToString();
                    cmd.Parameters.Add("@MiddleName", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][2].ToString();
                    cmd.Parameters.Add("@LastName", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][3].ToString();
                    cmd.Parameters.Add("@Company", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][4].ToString();
                    cmd.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = DtExcelData.Rows[i][5];
                    cmd.Parameters.Add("@CustomerStatusID", SqlDbType.Int).Value = DtExcelData.Rows[i][6];
                    cmd.Parameters.Add("@Email", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][7].ToString();
                    cmd.Parameters.Add("@Phone", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][8].ToString();
                    cmd.Parameters.Add("@MainAddress1", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][9].ToString();
                    cmd.Parameters.Add("@MainAddress2", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][10].ToString();
                    cmd.Parameters.Add("@MainAddress3", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][11].ToString();
                    cmd.Parameters.Add("@MainCity", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][12].ToString();
                    cmd.Parameters.Add("@MainState", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][13].ToString();
                    cmd.Parameters.Add("@MainZip", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][14].ToString();
                    cmd.Parameters.Add("@MainCountry", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][15].ToString();
                    cmd.Parameters.Add("@MailAddress1", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][16].ToString();
                    cmd.Parameters.Add("@MailAddress2", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][17].ToString();
                    cmd.Parameters.Add("@MailAddress3", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][18].ToString();
                    cmd.Parameters.Add("@MailCity", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][19].ToString();
                    cmd.Parameters.Add("@MailState", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][20].ToString();
                    cmd.Parameters.Add("@MailZip", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][21].ToString();
                    cmd.Parameters.Add("@MailCountry", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][22].ToString();
                    cmd.Parameters.Add("@CanLogin", SqlDbType.Bit).Value = DtExcelData.Rows[i][23];
                    cmd.Parameters.Add("@LoginName", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][24].ToString();
                    cmd.Parameters.Add("@BirthDate", SqlDbType.DateTime).Value = DtExcelData.Rows[i][25];
                    cmd.Parameters.Add("@CurrencyCode", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][26].ToString();
                    cmd.Parameters.Add("@LanguageID", SqlDbType.Int).Value = DtExcelData.Rows[i][27];
                    cmd.Parameters.Add("@Gender", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][28].ToString();
                    cmd.Parameters.Add("@TaxCode", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][29].ToString();
                    cmd.Parameters.Add("@TaxCodeTypeID", SqlDbType.Int).Value = DtExcelData.Rows[i][30];
                    cmd.Parameters.Add("@IsSalesTaxExempt", SqlDbType.Bit).Value = DtExcelData.Rows[i][31];
                    cmd.Parameters.Add("@SalesTaxCode", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][32].ToString();
                    cmd.Parameters.Add("@IsEmailSubscribed", SqlDbType.Bit).Value = DtExcelData.Rows[i][33];
                    cmd.Parameters.Add("@Notes", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][34].ToString();
                    cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime).Value = DtExcelData.Rows[i][35];
                    cmd.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = DtExcelData.Rows[i][36];
                    cmd.Parameters.Add("@CreatedBy", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][37].ToString();
                    cmd.Parameters.Add("@ModifiedBy", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][38].ToString();
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("Customer id=" + DtExcelData.Rows[i][0].ToString() + " " + " Updated");
                }
                else
                {
                    //sqlBulkCopy.DestinationTableName = "dbo.Customer";
                    //sqlBulkCopy.WriteToServer(DtExcelData.Rows[i].Table);

                    SqlCommand cmd = new SqlCommand("INSERT INTO Customer(FirstName,MiddleName,LastName,Company,CustomerTypeID,CustomerStatusID, Email, Phone, MainAddress1, MainAddress2, MainAddress3, MainCity, MainState, MainZip, MainCountry, MailAddress1, MailAddress2, MailAddress3, MailCity, MailState, MailZip, MailCountry,CanLogin,LoginName,BirthDate,CurrencyCode,LanguageID,Gender,TaxCode,TaxCodeTypeID,IsSalesTaxExempt,SalesTaxCode,IsEmailSubscribed,Notes,CreatedDate,ModifiedDate,CreatedBy,ModifiedBy) VALUES(@FirstName,@MiddleName,@LastName,@Company,@CustomerTypeID,@CustomerStatusID,@Email,@Phone,@MainAddress1,@MainAddress2,@MainAddress3,@MainCity,@MainState,@MainZip,@MainCountry,@MailAddress1,@MailAddress2,@MailAddress3,@MailCity,@MailState,@MailZip,@MailCountry,@CanLogin,@LoginName,@BirthDate,@CurrencyCode,@LanguageID,@Gender,@TaxCode,@TaxCodeTypeID,@IsSalesTaxExempt,@SalesTaxCode,@IsEmailSubscribed,@Notes,@CreatedDate,@ModifiedDate,@CreatedBy,@ModifiedBy)",Conn);
                    cmd.Parameters.Add("@FirstName", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][1].ToString();
                    cmd.Parameters.Add("@MiddleName", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][2].ToString();
                    cmd.Parameters.Add("@LastName", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][3].ToString();
                    cmd.Parameters.Add("@Company", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][4].ToString();
                    cmd.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = DtExcelData.Rows[i][5];
                    cmd.Parameters.Add("@CustomerStatusID", SqlDbType.Int).Value = DtExcelData.Rows[i][6];
                    cmd.Parameters.Add("@Email", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][7].ToString();
                    cmd.Parameters.Add("@Phone", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][8].ToString();
                    cmd.Parameters.Add("@MainAddress1", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][9].ToString();
                    cmd.Parameters.Add("@MainAddress2", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][10].ToString();
                    cmd.Parameters.Add("@MainAddress3", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][11].ToString();
                    cmd.Parameters.Add("@MainCity", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][12].ToString();
                    cmd.Parameters.Add("@MainState", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][13].ToString();
                    cmd.Parameters.Add("@MainZip", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][14].ToString();
                    cmd.Parameters.Add("@MainCountry", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][15].ToString();
                    cmd.Parameters.Add("@MailAddress1", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][16].ToString();
                    cmd.Parameters.Add("@MailAddress2", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][17].ToString();
                    cmd.Parameters.Add("@MailAddress3", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][18].ToString();
                    cmd.Parameters.Add("@MailCity", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][19].ToString();
                    cmd.Parameters.Add("@MailState", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][20].ToString();
                    cmd.Parameters.Add("@MailZip", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][21].ToString();
                    cmd.Parameters.Add("@MailCountry", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][22].ToString();
                    cmd.Parameters.Add("@CanLogin", SqlDbType.Bit).Value = DtExcelData.Rows[i][23];
                    cmd.Parameters.Add("@LoginName", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][24].ToString();
                    cmd.Parameters.Add("@BirthDate", SqlDbType.DateTime).Value = DtExcelData.Rows[i][25];
                    cmd.Parameters.Add("@CurrencyCode", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][26].ToString();
                    cmd.Parameters.Add("@LanguageID", SqlDbType.Int).Value = DtExcelData.Rows[i][27];
                    cmd.Parameters.Add("@Gender", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][28].ToString();
                    cmd.Parameters.Add("@TaxCode", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][29].ToString();
                    cmd.Parameters.Add("@TaxCodeTypeID", SqlDbType.Int).Value = DtExcelData.Rows[i][30];
                    cmd.Parameters.Add("@IsSalesTaxExempt", SqlDbType.Bit).Value = DtExcelData.Rows[i][31];
                    cmd.Parameters.Add("@SalesTaxCode", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][32].ToString();
                    cmd.Parameters.Add("@IsEmailSubscribed", SqlDbType.Bit).Value = DtExcelData.Rows[i][33];
                    cmd.Parameters.Add("@Notes", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][34].ToString();
                    cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime).Value = DtExcelData.Rows[i][35];
                    cmd.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = DtExcelData.Rows[i][36];
                    cmd.Parameters.Add("@CreatedBy", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][37].ToString();
                    cmd.Parameters.Add("@ModifiedBy", SqlDbType.NVarChar).Value = DtExcelData.Rows[i][38].ToString();
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("Customer id=" + DtExcelData.Rows[i][0].ToString() + " " + "Inserted");
                }
            }

        }
 
        public void orders()
        {
            Boolean flag = true;
            OleDbConnection exlconn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\kapil.sharma\\Desktop\\OrderTable.xlsx;Extended Properties='Excel 8.0;HDR=Yes'"); 
            OleDbCommand exlcommand_reader = new OleDbCommand("select * from [sheet1$]", exlconn);
            exlconn.Open();

            OleDbDataReader exl_dr = exlcommand_reader.ExecuteReader();
            while (exl_dr.Read())
            {               
                //SQL connection Object here:
                SqlConnection sqlconn = new SqlConnection(ssqlconnectionstring);
                SqlCommand sqlcommand_reader = new SqlCommand("select * from [dbo].[OrdersCopy]", sqlconn);
                sqlconn.Open();
                SqlDataReader sql_dr = sqlcommand_reader.ExecuteReader();
                while (sql_dr.Read())
                {
                    if (sql_dr[0].ToString() == exl_dr[0].ToString())
                    {
                        flag = true;                       
                        break;
                    }
                    else
                    {
                        flag = false;                        
                    }
                }
                sqlconn.Close();
                if (flag == true)
                {
                    //Updation
                    sqlconn = new SqlConnection(ssqlconnectionstring);                    
                    SqlCommand cmd = new SqlCommand("Update [dbo].[OrdersCopy] set OrderID=@OrderID, CustomerID=@CustomerID, OrderStatusID=@OrderStatusID, OrderDate=@OrderDate, CurrencyCode=@CurrencyCode, WarehouseID=@WarehouseID, ShipMethodID=@ShipMethodID, OrderTypeID=@OrderTypeID, PriceTypeID=@PriceTypeID, FirstName=@FirstName, MiddleName=@MiddleName, LastName=@LastName, NameSuffix=@NameSuffix, Company=@Company, Address1=@Address1, Address2=@Address2, Address3=@Address3, City=@City, State=@State, Zip=@Zip, Country=@Country, County=@County, Email=@Email, Phone=@Phone, Notes=@Notes, Total=@Total, SubTotal=@SubTotal, TaxTotal=@TaxTotal, ShippingTotal=@ShippingTotal, DiscountTotal=@DiscountTotal, DiscountPercent=@DiscountPercent, WeightTotal=@WeightTotal, CreatedDate=@CreatedDate, ModifiedDate=@ModifiedDate, CreatedBy=@CreatedBy, ModifiedBy=@ModifiedBy where OrderID="+exl_dr[0].ToString());
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = sqlconn;
                    //Values into parameters
                    cmd.Parameters.AddWithValue("@OrderID", Int32.Parse(exl_dr[0].ToString()));
                    cmd.Parameters.AddWithValue("@CustomerID", Int32.Parse(exl_dr[1].ToString()));
                    cmd.Parameters.AddWithValue("@OrderStatusID", Int32.Parse(exl_dr[2].ToString()));
                    cmd.Parameters.AddWithValue("@OrderDate", exl_dr[3].ToString());
                    cmd.Parameters.AddWithValue("@CurrencyCode", exl_dr[4].ToString());
                    cmd.Parameters.AddWithValue("@WarehouseID", Int32.Parse(exl_dr[5].ToString()));
                    cmd.Parameters.AddWithValue("@ShipMethodID", Int32.Parse(exl_dr[6].ToString()));
                    cmd.Parameters.AddWithValue("@OrderTypeID", Int32.Parse(exl_dr[7].ToString()));
                    cmd.Parameters.AddWithValue("@PriceTypeID", Int32.Parse(exl_dr[8].ToString()));
                    cmd.Parameters.AddWithValue("@FirstName", exl_dr[9].ToString());
                    cmd.Parameters.AddWithValue("@MiddleName", exl_dr[10].ToString());
                    cmd.Parameters.AddWithValue("@LastName", exl_dr[11].ToString());
                    cmd.Parameters.AddWithValue("@NameSuffix", exl_dr[12].ToString());
                    cmd.Parameters.AddWithValue("@Company", exl_dr[13].ToString());
                    cmd.Parameters.AddWithValue("@Address1", exl_dr[14].ToString());
                    cmd.Parameters.AddWithValue("@Address2", exl_dr[15].ToString());
                    cmd.Parameters.AddWithValue("@Address3", exl_dr[16].ToString());
                    cmd.Parameters.AddWithValue("@City", exl_dr[17].ToString());
                    cmd.Parameters.AddWithValue("@State", exl_dr[18].ToString());
                    cmd.Parameters.AddWithValue("@Zip", exl_dr[19].ToString());
                    cmd.Parameters.AddWithValue("@Country", exl_dr[20].ToString());
                    cmd.Parameters.AddWithValue("@County", exl_dr[21].ToString());
                    cmd.Parameters.AddWithValue("@Email", exl_dr[22].ToString());
                    cmd.Parameters.AddWithValue("@Phone", exl_dr[23].ToString());
                    cmd.Parameters.AddWithValue("@Notes", exl_dr[24].ToString());
                    cmd.Parameters.AddWithValue("@Total", Decimal.Parse(exl_dr[25].ToString()));
                    cmd.Parameters.AddWithValue("@SubTotal", Decimal.Parse(exl_dr[26].ToString()));
                    cmd.Parameters.AddWithValue("@TaxTotal", Decimal.Parse(exl_dr[27].ToString()));
                    cmd.Parameters.AddWithValue("@ShippingTotal", Decimal.Parse(exl_dr[28].ToString()));
                    cmd.Parameters.AddWithValue("@DiscountTotal", Decimal.Parse(exl_dr[29].ToString()));
                    cmd.Parameters.AddWithValue("@DiscountPercent", Decimal.Parse(exl_dr[30].ToString()));
                    cmd.Parameters.AddWithValue("@WeightTotal", Decimal.Parse(exl_dr[31].ToString()));
                    cmd.Parameters.AddWithValue("@CreatedDate", exl_dr[32].ToString());
                    cmd.Parameters.AddWithValue("@ModifiedDate", exl_dr[33].ToString());
                    cmd.Parameters.AddWithValue("@CreatedBy", exl_dr[34].ToString());
                    cmd.Parameters.AddWithValue("@ModifiedBy", exl_dr[35].ToString());
                    sqlconn.Open();
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("OrderID: " + exl_dr[0] + " Updated");

                }
                else
                {
                    //Insertion
                    sqlconn = new SqlConnection(ssqlconnectionstring);                    
                    SqlCommand cmd = new SqlCommand("Insert into [dbo].[OrdersCopy] values (@OrderID, @CustomerID, @OrderStatusID, @OrderDate, @CurrencyCode, @WarehouseID, @ShipMethodID, @OrderTypeID, @PriceTypeID, @FirstName, @MiddleName, @LastName, @NameSuffix, @Company, @Address1, @Address2, @Address3, @City, @State, @Zip, @Country, @County, @Email, @Phone, @Notes, @Total, @SubTotal, @TaxTotal, @ShippingTotal, @DiscountTotal, @DiscountPercent, @WeightTotal, @CreatedDate, @ModifiedDate, @CreatedBy, @ModifiedBy)");
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = sqlconn;                   
                    cmd.Parameters.AddWithValue("@OrderID", Int32.Parse(exl_dr[0].ToString()));
                    cmd.Parameters.AddWithValue("@CustomerID", Int32.Parse(exl_dr[1].ToString()));
                    cmd.Parameters.AddWithValue("@OrderStatusID", Int32.Parse(exl_dr[2].ToString()));                   
                    cmd.Parameters.AddWithValue("@OrderDate", exl_dr[3].ToString());
                    cmd.Parameters.AddWithValue("@CurrencyCode", exl_dr[4].ToString());
                    cmd.Parameters.AddWithValue("@WarehouseID", Int32.Parse(exl_dr[5].ToString()));
                    cmd.Parameters.AddWithValue("@ShipMethodID", Int32.Parse(exl_dr[6].ToString()));
                    cmd.Parameters.AddWithValue("@OrderTypeID", Int32.Parse(exl_dr[7].ToString()));
                    cmd.Parameters.AddWithValue("@PriceTypeID", Int32.Parse(exl_dr[8].ToString()));
                    cmd.Parameters.AddWithValue("@FirstName", exl_dr[9].ToString());
                    cmd.Parameters.AddWithValue("@MiddleName", exl_dr[10].ToString());
                    cmd.Parameters.AddWithValue("@LastName", exl_dr[11].ToString());
                    cmd.Parameters.AddWithValue("@NameSuffix", exl_dr[12].ToString());
                    cmd.Parameters.AddWithValue("@Company", exl_dr[13].ToString());
                    cmd.Parameters.AddWithValue("@Address1", exl_dr[14].ToString());
                    cmd.Parameters.AddWithValue("@Address2", exl_dr[15].ToString());
                    cmd.Parameters.AddWithValue("@Address3", exl_dr[16].ToString());
                    cmd.Parameters.AddWithValue("@City", exl_dr[17].ToString());
                    cmd.Parameters.AddWithValue("@State", exl_dr[18].ToString());
                    cmd.Parameters.AddWithValue("@Zip", exl_dr[19].ToString());
                    cmd.Parameters.AddWithValue("@Country", exl_dr[20].ToString());
                    cmd.Parameters.AddWithValue("@County", exl_dr[21].ToString());
                    cmd.Parameters.AddWithValue("@Email", exl_dr[22].ToString());
                    cmd.Parameters.AddWithValue("@Phone", exl_dr[23].ToString());
                    cmd.Parameters.AddWithValue("@Notes", exl_dr[24].ToString());
                    cmd.Parameters.AddWithValue("@Total", Decimal.Parse(exl_dr[25].ToString()));
                    cmd.Parameters.AddWithValue("@SubTotal", Decimal.Parse(exl_dr[26].ToString()));
                    cmd.Parameters.AddWithValue("@TaxTotal", Decimal.Parse(exl_dr[27].ToString()));
                    cmd.Parameters.AddWithValue("@ShippingTotal", Decimal.Parse(exl_dr[28].ToString()));
                    cmd.Parameters.AddWithValue("@DiscountTotal", Decimal.Parse(exl_dr[29].ToString()));
                    cmd.Parameters.AddWithValue("@DiscountPercent", Decimal.Parse(exl_dr[30].ToString()));
                    cmd.Parameters.AddWithValue("@WeightTotal", Decimal.Parse(exl_dr[31].ToString()));                    
                    cmd.Parameters.AddWithValue("@CreatedDate", exl_dr[32].ToString());                    
                    cmd.Parameters.AddWithValue("@ModifiedDate", exl_dr[33].ToString());
                    cmd.Parameters.AddWithValue("@CreatedBy", exl_dr[34].ToString());
                    cmd.Parameters.AddWithValue("@ModifiedBy", exl_dr[35].ToString());
                    sqlconn.Open();
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("OrderID: " + exl_dr[0] + "Inserted");
                }
                sqlconn.Close();
            }
        }
        public void order_status()
        {
            int flag = 0;
            OleDbConnection exlconn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\kapil.sharma\\Desktop\\OrderStatus.xlsx;Extended Properties='Excel 8.0;HDR=Yes'");
            OleDbCommand exlcommand_reader = new OleDbCommand("select * from [sheet1$]", exlconn);
            exlconn.Open();
            OleDbDataReader exl_dr = exlcommand_reader.ExecuteReader();
            while (exl_dr.Read())
            {
                //SQL connection Object here:
                SqlConnection sqlconn = new SqlConnection(ssqlconnectionstring);
                SqlCommand sqlcommand_reader = new SqlCommand("select * from [dbo].[OrderStatus]", sqlconn);
                sqlconn.Open();
                SqlDataReader sql_dr = sqlcommand_reader.ExecuteReader();
                while (sql_dr.Read())
                {
                    if (sql_dr[0].ToString() == exl_dr[0].ToString())
                    {
                        if (sql_dr[1].ToString() != exl_dr[1].ToString())
                        {
                            flag = 1;
                        }
                        else
                        {
                            flag = 0;
                        }
                        break;
                    }
                    else
                    {
                        flag = 2;
                    }
                }
                sqlconn.Close();
                if (flag==1)
                {
                    //Updating record w/o overwriting existing OrderID
                    sqlconn = new SqlConnection(ssqlconnectionstring);
                    SqlCommand cmd = new SqlCommand("Insert into [dbo].[OrderStatus] values (@OrderID, @OrderStatusID)");
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = sqlconn;
                    cmd.Parameters.AddWithValue("@OrderID", Int32.Parse(exl_dr[0].ToString()));
                    cmd.Parameters.AddWithValue("@OrderStatusID", Int32.Parse(exl_dr[1].ToString()));
                    sqlconn.Open();
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("OrderID: " + exl_dr[0] + " Updated");
                }
                else if (flag==2)
                {
                    //Inserting new OrderID
                    sqlconn = new SqlConnection(ssqlconnectionstring);
                    SqlCommand cmd = new SqlCommand("Insert into [dbo].[OrderStatus] values (@OrderID, @OrderStatusID)");
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = sqlconn;
                    cmd.Parameters.AddWithValue("@OrderID", Int32.Parse(exl_dr[0].ToString()));
                    cmd.Parameters.AddWithValue("@OrderStatusID", Int32.Parse(exl_dr[1].ToString()));
                    sqlconn.Open();
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("OrderID: " + exl_dr[0] + " Inserted");
                }
                sqlconn.Close();
            }
        }
        public static void Main(string[] args)
        {
            Class1 obj = new Class1();
        obj.importdatafromexcel();           
            char choice;
            int option;     
            do
            {
                Console.WriteLine("Choose the operation you wanna do:\n1.Customer\n2.Order\n3.Order Status");
                option = Convert.ToInt32((Console.ReadLine()));
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
                    break;
                }
                Console.Write("Do you wish to continue: ");
                choice = Console.ReadKey().KeyChar;

            } while (choice == 'Y' || choice == 'y');

            Console.ReadLine();
        }

    }
}
