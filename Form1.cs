using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;


namespace Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string str;
            int rCnt;
            int cCnt;

            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Excel (*.XLS, *.XLSX)|*.XLS;*.XLSX";
            opf.ShowDialog();
            System.Data.DataTable tb = new System.Data.DataTable();
            string filename = opf.FileName;

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExcelRange;

            ExcelWorkBook = ExcelApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false,
                false, 0, true, 1, 0);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            ExcelRange = ExcelWorkSheet.UsedRange;
            for (rCnt = 1; rCnt <= ExcelRange.Rows.Count; rCnt++)
            {
                dataGridViewStudent.Rows.Add(1);
                for (cCnt = 1; cCnt <= 3; cCnt++)
                {
                    str = (string)(ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                    dataGridViewStudent.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                }
            }
            ExcelWorkBook.Close(true, null, null);
            ExcelApp.Quit();

            releaseObject(ExcelWorkSheet);
            releaseObject(ExcelWorkBook);
            releaseObject(ExcelApp);

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
                MessageBox.Show("Unable to release the object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridViewStudent.Sort(dataGridViewStudent.Columns["Given"], ListSortDirection.Descending);
            dataGridViewStudent.Sort(dataGridViewStudent.Columns["Family"], ListSortDirection.Descending);
            dataGridViewStudent.Sort(dataGridViewStudent.Columns["Pat"], ListSortDirection.Descending);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridViewStudent.Rows.Count - 1; i++)
            {
                bool isVisible = false;
                for (int j = 0; j < dataGridViewStudent.Columns.Count; j++)
                {
                    if (dataGridViewStudent[j, i].Value.ToString() == textBox1.Text)
                    {
                        isVisible = true;
                    }
                }
                dataGridViewStudent.Rows[i].Visible = isVisible;
            }


        }

        private void button5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridViewStudent.Rows.Count - 1; i++)

            {
                dataGridViewStudent.Rows[i].Visible = true;
            }
        }
    }
}

