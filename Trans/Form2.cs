using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Sql;
using Excel = Microsoft.Office.Interop;

namespace Trans
{
    public partial class Form2 : Form
    {
        int n = 1;
        public Form2()
        {
            InitializeComponent();
        }
        string ExcelFolderPath = @"C:\Users\Vinh02\Downloads\214.xlsx";
        private void Form2_Load(object sender, EventArgs e)
        {
            OpenExcelFile();
        }
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case 0x84:
                    base.WndProc(ref m);
                    if ((int)m.Result == 0x1)
                        m.Result = (IntPtr)0x2;
                    return;
            }

            base.WndProc(ref m);
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void OpenExcelFile()
        {
            System.Data.DataTable NewDT;

            foreach (var file in Directory.GetFiles(this.ExcelFolderPath).Where(p => p.Contains("Contacts")))
            {
                Microsoft.Office.Interop.Excel.Application excelApp = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
                Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = null;
                Microsoft.Office.Interop.Excel.Range excelRange = null;
                try
                {
                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelWorkbook = excelApp.Workbooks.Open(file);
                    excelWorksheet = excelWorkbook.Sheets[1];
                    excelRange = excelWorksheet.UsedRange;
                    int rowCount = excelRange.Rows.Count;
                    int colCount = excelRange.Columns.Count;
                    object[,] data = (object[,])excelRange.Cells.Value;
                    NewDT = new System.Data.DataTable();
                    for (int i = 1; i <= colCount; i++)
                    {
                        if (data[1, i] != null)
                        {
                            NewDT.Columns.Add(data[1, i].ToString(), typeof(string));
                        }
                        else
                        {
                            NewDT.Columns.Add("", typeof(string));
                        }
                    }
                    DataRow curRow = null;
                    for (int i = 2; i <= rowCount; i++)
                    {
                        curRow = NewDT.NewRow();
                        for (int j = 1; j <= colCount; j++)
                        {
                            if (data[i, j] != null)
                            {
                                curRow[j - 1] = data[i, j].ToString();
                            }
                        }
                        NewDT.Rows.Add(curRow);
                    }
                    dataGridView1.DataSource = NewDT;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    //GC.Collect();
                    //GC.WaitForPendingFinalizers();
                    if (excelRange != null)
                    {
                        Marshal.ReleaseComObject(excelRange);
                    }
                    if (excelWorksheet != null)
                    {
                        Marshal.ReleaseComObject(excelWorksheet);
                    }
                    if (excelWorkbook != null)
                    {
                        excelWorkbook.Close();
                        Marshal.ReleaseComObject(excelWorkbook);
                    }
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
            }
        }//end function

        public void CreateDataTableForExcelData(String FileName)
        {
            OleDbConnection ExcelConnection = null;
            string filePath = @"C:\Users\Vinh02\Downloads\214.xlsx";
            System.Data.DataTable dtNew = new System.Data.DataTable();
            string strExt = "";
            strExt = FileName.Substring(FileName.LastIndexOf("."));
            if (strExt == ".xls")
            {
                ExcelConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + hdnFileName.Value + ";Extended Properties=Excel 8.0;");
            }
            else
            {
                if (strExt == ".xlsx")
                {
                    ExcelConnection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + hdnFileName.Value + ";Extended Properties=Excel 12.0;");
                }
            }
            try
            {
                ExcelConnection.Open();
                System.Data.DataTable dt = ExcelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                OleDbCommand ExcelCommand = new OleDbCommand(@"SELECT * FROM [" + ddlTableName.SelectedValue + @"]", ExcelConnection);
                OleDbDataAdapter ExcelAdapter = new OleDbDataAdapter(ExcelCommand);
                DataSet ExcelDataSet = new DataSet();
                ExcelAdapter.Fill(dtExcel);
                ExcelConnection.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }


    }
}
