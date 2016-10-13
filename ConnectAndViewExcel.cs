using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace Training_Orders_Engine
{
    class ConnectAndViewExcel
    {
        DataTable DtExcelData;
        DataTable DtSqlData;
        OleDbConnection ExcelConn;
        OleDbDataAdapter DaExcelcmd;
        SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(Connection.Conn);

        public void excel()
        {
            string Exfilepath = @"C:\Users\siddharth.bhatnagar\Desktop\Customer.xlsx";
            DtExcelData = new DataTable();
            DtSqlData = new DataTable();
  

            //If you MS Excel 2007 then use below lin instead of above line
            ExcelConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Exfilepath + "';Extended Properties='Excel 12.0;hdr=yes;'");

            DaExcelcmd = new OleDbDataAdapter("select * from [Sheet1$]", ExcelConn);
            ExcelConn.Open();
            DaExcelcmd.Fill(DtExcelData);

            SqlCommand com = new SqlCommand("select * from Customer", Connection.Conn);
            com.CommandType = CommandType.Text;
            SqlDataAdapter DaSqlcmd = new SqlDataAdapter(com);
            DaSqlcmd.Fill(DtSqlData);

            //sqlBulkCopy.DestinationTableName = "dbo.Customer";
            //sqlBulkCopy.WriteToServer(DtExcelData);
            customers();
            
        }

        public void customers()
        {

            int sqlrow=DtSqlData.Rows.Count;
            bool Flag=false;
            int excelrow = DtExcelData.Rows.Count;

            for(int i=0;i<excelrow;i++)
            {
                for(int j=0;j<sqlrow;j++)
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
                if(Flag)
                {
                    SqlCommand cmd = new SqlCommand("UPDATE Customer SET CustomerID=@CustomerID,FirstName=@FirstName,MiddleName=@MiddleName, LastName=@LastName, Company=@Company, CustomerTypeID=@CustomerTypeID, CustomerStatusID=@CustomerStatusID, Email=@Email, Phone=@Phone, MainAddress1=@MainAddress1, MainAddress2=@MainAddress2, MainAddress3=@MainAddress3, MainCity=@MainCity, MainState=@MainState, MainZip=@MainZip, MainCountry=@MainCountry, MailAddress1=@MailAddress1, MailAddress2=@MailAddress2, MailAddress3=@MailAddress3, MailCity=@MailCity, MailState=@MailState, MailZip=@MailZip, MailCountry=@MailCountry,CanLogin=@CanLogin, LoginName=@LoginName, BirthDate=@BirthDate, CurrencyCode=@CurrencyCode, LanguageID=@LanguageID, Gender=@Gender, TaxCode=@TaxCode, TaxCodeTypeID=@TaxCodeTypeID, IsSalesTaxExempt=@IsSalesTaxExempt, SalesTaxCode=@SalesTaxCode, IsEmailSubscribed=@IsEmailSubscribed, Notes=@Notes, CreatedDate=@CreatedDate, ModifiedDate=@ModifiedDate, CreatedBy=@CreatedBy, ModifiedBy=@ModifiedBy WHERE CustomerID=@CustomerID", Connection.Conn);
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
                }
                else
                {
                    Console.WriteLine("inserted new record");
                   //sqlBulkCopy.DestinationTableName = "dbo.Customer";
                   //sqlBulkCopy.WriteToServer(DtExcelData.Rows[i].Table);

                    SqlCommand cmd = new SqlCommand("INSERT INTO Customer(CustomerID,FirstName,MiddleName,LastName,Company,CustomerTypeID,CustomerStatusID, Email, Phone, MainAddress1, MainAddress2, MainAddress3, MainCity, MainState, MainZip, MainCountry, MailAddress1, MailAddress2, MailAddress3, MailCity, MailState, MailZip, MailCountry,CanLogin,LoginName,BirthDate,CurrencyCode,LanguageID,Gender,TaxCode,TaxCodeTypeID,IsSalesTaxExempt,SalesTaxCode,IsEmailSubscribed,Notes,CreatedDate,ModifiedDate,CreatedBy,ModifiedBy) VALUES(@CustomerID,@FirstName,@MiddleName,@LastName,@Company,@CustomerTypeID,@CustomerStatusID,@Email,@Phone,@MainAddress1,@MainAddress2,@MainAddress3,@MainCity,@MainState,@MainZip,@MainCountry,@MailAddress1,@MailAddress2,@MailAddress3,@MailCity,@MailState,@MailZip,@MailCountry,@CanLogin,@LoginName,@BirthDate,@CurrencyCode,@LanguageID,@Gender,@TaxCode,@TaxCodeTypeID,@IsSalesTaxExempt,@SalesTaxCode,@IsEmailSubscribed,@Notes,@CreatedDate,@ModifiedDate,@CreatedBy,@ModifiedBy)", Connection.Conn);
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
                    cmd.Parameters.Add("@ModifiedBy",SqlDbType.NVarChar).Value = DtExcelData.Rows[i][38].ToString();
                    cmd.ExecuteNonQuery();
                }

            }
          
        }
    }
}
