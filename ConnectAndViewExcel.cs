using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Training_Orders_Engine
{
    class ConnectAndViewExcel
    {
        DataTable DtExcelData;
        DataTable DtSqlData;
        public void excel()
        {
            string Exfilepath = @"C:\Users\siddharth.bhatnagar\Desktop\Customer.xlsx";
            DtExcelData = new DataTable();
            DtSqlData = new DataTable();
  
            OleDbConnection ExcelConn;
            OleDbDataAdapter DaExcelcmd;
            OleDbDataAdapter DaSqlcmd;
            SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(Connection.Conn);

            //If you MS Excel 2007 then use below lin instead of above line
            ExcelConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Exfilepath + "';Extended Properties='Excel 12.0;hdr=yes;'");

            DaExcelcmd = new OleDbDataAdapter("select * from [Sheet1$]", ExcelConn);
            ExcelConn.Open();
            DaExcelcmd.Fill(DtExcelData);

            DaSqlcmd = new OleDbDataAdapter("select * from Customer",Connection.Conn.ConnectionString);
            DaSqlcmd.Fill(DtSqlData);



            sqlBulkCopy.DestinationTableName = "dbo.Customer";
            sqlBulkCopy.WriteToServer(DtExcelData);
            ExcelConn.Close();
        }

        public void customers()
        {
            foreach (DataRow dr in DtExcelData.Rows)
            {
                foreach (DataRow dr1 in DtSqlData.Rows)
                {
                    if (dr[0]==dr1[0])
                    {
                        dr[dc] = null;
                    }
                }

            }
        }
    }
}
