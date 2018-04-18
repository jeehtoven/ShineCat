using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using OfficeOpenXml;

namespace ShineCat
{
    class Program
    {
        static void Main(string[] args)
        {
            bool result = CreateFile();
            string filePath = "ShineCat.xls";

            if (result == true) System.Diagnostics.Process.Start(filePath);
        }

        static public bool CreateFile()
        {
            DataSet ds = CreateData();
            var sourceFile = new FileInfo("ShineCat.xls");

            using (ExcelPackage pck = new ExcelPackage())
            {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Accounts");
                    ws.Cells["A1"].LoadFromDataTable(ds.Tables[0], true);
                    pck.SaveAs(sourceFile);
            }

            return true;

        }

        static public DataSet CreateData()
        {
            // Create two DataTable instances.
            DataTable table1 = new DataTable("patients");
            table1.Columns.Add("name");
            table1.Columns.Add("id");
            table1.Rows.Add("sam", 1);
            table1.Rows.Add("mark", 2);

            DataTable table2 = new DataTable("medications");
            table2.Columns.Add("id");
            table2.Columns.Add("medication");
            table2.Rows.Add(1, "atenolol");
            table2.Rows.Add(2, "amoxicillin");

            // Create a DataSet and put both tables in it.
            DataSet set = new DataSet("office");
            set.Tables.Add(table1);
            set.Tables.Add(table2);

            // Visualize DataSet.
            Console.WriteLine(set.GetXml());

            return set;
        }
    }
}
