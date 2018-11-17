using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelFile = Microsoft.Office.Interop.Excel;
using System.Linq;

namespace DataStockRender
{
    public partial class Form1 : Form
    {
        string filePath = string.Empty;
        CheckedListBox checkedList = new CheckedListBox();
        public Form1()
        {
            InitializeComponent();
            //CheckBox box;
            //for (int i = 0; i < 20; i++)
            //{
            //    box = new CheckBox();
            //    box.Tag = i.ToString();
            //    box.Text = "aaaaaaaa" + i*100;
         
            //    box.SetBounds(0, 0, 210 , 20);
            //    flowLayoutPanel1.Controls.Add(box);
            //}

        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
        
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        protected override CreateParams CreateParams
        {
            get
            {
                const int CS_DROPSHADOW = 0x20000;
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_DROPSHADOW;
                return cp;
            }

        }
        
        private void btnSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            fileDialog.ShowDialog();
            
            filePath = fileDialog.FileName;
            txtFilePath.Enabled = false;
            txtFilePath.Text = filePath;
            //for (int i = 1; i <= rowCount; i++)
            //{
            //    for (int j = 1; j <= colCount; j++)
            //    {
            //        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
            //        {
            //            Console.WriteLine(xlRange.Cells[i, j].Value);
            //            //dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
            //        }
            //        // Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");  

            //        //add useful things here!     
            //    }
            //}
        }

        private void btnReadExcel_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelFile.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelFile.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                ExcelFile._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                ExcelFile.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                
                // create checkbox
                for (int j = 1; j <= colCount; j++)
                {
                    CheckBox box = new CheckBox();
                    var label = xlRange.Cells[1, j].Value;
                    box.Text = label;
                    box.SetBounds(0, 0, 210, 20);
                    flowLayoutPanel1.Controls.Add(box);
                    //checkedList.Items.Add(box);
                }

                string[] label_default = {"DeltaPricePast","DeltaKLMB","DeltaKLMBKL",
                    "DeltaP/E", "DeltaP/S", "DeltaP/B" };
                for(int j = 0; j < label_default.Length; j++)
                {
                    CheckBox box = new CheckBox();
                    box.Text = label_default[j];
                    flowLayoutPanel1.Controls.Add(box);
                    //checkedList.Items.Add(box);
                }
                
                #region Release object
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                #endregion
            }
            catch
            {
                // exception
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            // count number of selected attribute
            var checkedBoxes = flowLayoutPanel1.Controls.OfType<CheckBox>().Count(c => c.Checked);
            Console.WriteLine(checkedBoxes);

        }
    }
}
