using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;

namespace ExcelParser
{
    public class Parser
    {
        public DataTable ParseFunction(string filepath, string filetype)
        {
            string sheetName = "";
            if (filetype == "eq") sheetName = "Equities";
            else if (filetype == "cb") sheetName = "Bonds";

            OleDbConnection con;
            con = new OleDbConnection(@"provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + filepath + "';Extended Properties='Excel 12.0;IMEX=1'");
            DataSet ds;
            OleDbDataAdapter adapter;
            DataTable dt = new DataTable();
            try
            {
                adapter = new OleDbDataAdapter("select * from [" + sheetName + "$]", con);
                adapter.TableMappings.Add("Table", "TestTable");
                ds = new DataSet();
                adapter.Fill(ds);
                dt = ds.Tables[0].Rows.Cast<DataRow>().Where(row => !row.ItemArray.All(field => field is System.DBNull || string.Compare((field as string).Trim(), string.Empty) == 0)).CopyToDataTable();
            }
            catch (Exception ex) { Debug.WriteLine(ex.Message); }
            finally { con.Close(); }
            return dt;
        }
    }
}
