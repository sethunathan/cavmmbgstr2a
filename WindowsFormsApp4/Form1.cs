using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
namespace WindowsFormsApp4
{
    public static class StringExtension
    {
        public static string GetLast(this string source, int tail_length)
        {
            if (tail_length >= source.Length)
                return source;
            return source.Substring(source.Length - tail_length);
        }
    }
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;
            
            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if(xlApp == null)
            {
                MessageBox.Show("Excel Application not Installed");
                return;
            }
            else
            {
                
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "ID";
                xlWorkSheet.Cells[1, 2] = "Name";

                xlWorkSheet.Cells[2, 1] = "1";
                xlWorkSheet.Cells[2, 2] = "Seasdthu";


                xlWorkBook.SaveAs("d:\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true,misValue,misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                MessageBox.Show("File Created ");
                 
            }
            
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;int rCnt;int cCnt;int rw = 0;int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"d:\csharp-Excel.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            for(rCnt=1;rCnt <=rw;rCnt++)
            {
                for(cCnt=1;cCnt<=cl;cCnt++)
                {
                    str = (string)(range.Cells[rCnt,cCnt] as Excel.Range).Value2;
                    MessageBox.Show(str);
                }
            }
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlApp);


        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.DisplayAlerts=false;
            string filePath = @"d:\csharp-Excel.xlsx";

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Sheets xlWorkSheets = xlWorkBook.Worksheets;

            var xlNewSheet = (Excel.Worksheet)xlWorkSheets.Add(xlWorkSheets[1], Type.Missing, Type.Missing, Type.Missing);
            xlNewSheet.Name = "Welcome new";

            xlNewSheet.Cells[1, 1] = "Welcome to New Sheet";
            xlNewSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            xlNewSheet.Select();
            xlWorkBook.Save();
            xlWorkBook.Close();

            releaseObject(xlNewSheet);
            releaseObject(xlWorkSheets);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Sheet Created");

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        

        private void Button4_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.OleDb.OleDbConnection yaalconnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;

            // public static string path = @"C:\src\RedirectApplication\RedirectApplication\301s.xlsx";
            // public static string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
            
                yaalconnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='Z:\\33BRKPR0174G1ZS_032019_R2A.xlsx';Extended Properties=Excel 12.0;");
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [B2B$]",yaalconnection);


               // Provider = Microsoft.Jet.OLEDB.4.0; Data Source = E:\33BRKPR0174G1ZS_032019_R2A.xlsx

                //Provider = Microsoft.ACE.OLEDB.12.0; Data Source = D:\json\test.accdb


                MyCommand.TableMappings.Add("Table", "TestTable");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                dataGridView1.DataSource = DtSet.Tables[0];

                yaalconnection.Close();


                System.Data.OleDb.OleDbConnection con = new System.Data.OleDb.OleDbConnection(" Provider = Microsoft.ACE.OLEDB.12.0; Data Source =D:\\json\\test.accdb");

                string StrQuery ;
                System.Data.OleDb.OleDbCommand comm = con.CreateCommand();
                con.Open();

                String gstno;
                string mystring;
                mystring = "";
                string res;
                res = "";

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value == DBNull.Value)
                        {
                            dataGridView1.Rows[i].Cells[j].Value =0;
                        }
                    }
                }
                dataGridView1.Update();
               

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[2].Value != null && dataGridView1.Rows[i].Cells[2].ToString() != String.Empty)
                    {
                        string value = (string)dataGridView1.Rows[i].Cells[2].Value;
                        mystring = value;
                        res = mystring.GetLast(5);
                    }
                  
                    if (res == "Total")
                    { 
                       StrQuery = @"INSERT INTO test(gstno,amount) VALUES ('" + dataGridView1.Rows[i].Cells[0].Value + "'," + dataGridView1.Rows[i].Cells[5].Value + ");";

                       //string StrQuery = "INSERT INTO tableName VALUES ('" + dataGridView1.Rows[i].Cells[0].Value + "',' " + dataGridView1.Rows[i].Cells[1].Value + "', '" + dataGridView1.Rows[i].Cells[2].Value + "', '" + dataGridView1.Rows[i].Cells[3].Value + "',' " + dataGridView1.Rows[i].Cells[4].Value + "')";

                    // StrQuery = @"INSERT INTO test VALUES (" + dataGridView1.Rows[i].Cells["ColumnName"].Text + ", "
                    //  + dataGridView1.Rows[i].Cells["ColumnName"].Text + ");";

                      comm.CommandText = StrQuery;
                      comm.ExecuteNonQuery();
                }

                }
                //  cmd.CommandText = "Insert into test(test)Values(1)";
                comm.Connection = con;
                comm.ExecuteNonQuery();
                MessageBox.Show("Record Submitted", "Congrats");
                con.Close();

                MessageBox.Show("Completed");

            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.Message);
            }
        }
    }
}
