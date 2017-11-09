using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Ex = Microsoft.Office.Interop.Excel;

namespace ScrapAnalyzer
{
    public partial class Summary : Form
    {
        private Dictionary<string, Dictionary<string, double>> info;
        private string dataType;

        public Summary(Dictionary<string, Dictionary<string, double>> _info, string _dataType)
        {
            InitializeComponent();
            info = _info;
            dataType = _dataType;
        }

        private void fillData()
        {
            string productName;
            Dictionary<string, double>.KeyCollection productScraps;
            Dictionary<string, double> productInfo;
            data.Columns.Add("Product name", "Product name");
            int index;
            foreach (KeyValuePair<string, Dictionary<string, double>> product in info)
            {
                productName = product.Key;
                productScraps = product.Value.Keys;
                productInfo = product.Value;
                index = data.Rows.Add();
                data.Rows[index].Cells["Product name"].Value = productName;
                foreach (string scrapName in productScraps)
                {
                    if (!data.Columns.Contains(scrapName))
                    {
                        data.Columns.Add(scrapName, scrapName);
                    }
                    if (data.Rows[index].Cells[scrapName].Value == null)
                    {
                        data.Rows[index].Cells[scrapName].Value = 0;
                    }
                    double scrapValue = double.Parse(data.Rows[index].Cells[scrapName].Value.ToString());
                    data.Rows[index].Cells[scrapName].Value = scrapValue + productInfo[scrapName];
                }
            }

            foreach (DataGridViewRow row in data.Rows)
            {
                for (int i = 0; i < data.Columns.Count; i++)
                {
                    if (row.Cells[i].Value == null)
                    {
                        row.Cells[i].Value = 0;
                    }
                }
            }
        }

        private void Summary_SizeChanged(object sender, EventArgs e)
        {
            data.Size = new Size(this.Size.Width - 16, this.Size.Height - 39);
        }

        private void Summary_Load(object sender, EventArgs e)
        {
            fillData();
            addSumToTable();
            this.MaximumSize = data.PreferredSize + new Size(0, 7);
        }

        private void addSumToTable()
        {
            string columnName = "Sum of broken " + dataType;
            data.Columns.Add(columnName, columnName);
            double sum;
            foreach (DataGridViewRow row in data.Rows)
            {
                sum = 0;
                for (int i = 1; i < row.Cells.Count - 1; i++)
                {
                    sum += Double.Parse(row.Cells[i].Value.ToString());
                }
                row.Cells[columnName].Value = sum;
            }
            ////////////////////////////////////////////////////////////
            int newRowIndex = data.Rows.Add("Total");
            string currColumn;
            for (int i = 1; i < data.Columns.Count; i++)
            {
                currColumn = data.Columns[i].HeaderText;
                sum = 0;
                for (int j = 0; j < data.Rows.Count - 1; j++)
                {
                    //MessageBox.Show(Double.Parse(data.Rows[j].Cells[currColumn].Value.ToString()).ToString() + "Row: " + j.ToString()+"\nColumn: "+currColumn);
                    sum += Double.Parse(data.Rows[j].Cells[currColumn].Value.ToString());
                }
                data.Rows[newRowIndex].Cells[currColumn].Value = sum;
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.S))
            {
                //Open a save file dialog
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb";
                sfd.FileName = "export.xls";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    exportData(data, sfd.FileName);
                }
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void exportData(DataGridView d, string filename)
        {
            /** string stOutput = "";
            // Export titles:
            string sHeaders = "";

            for (int j = 0; j < d.Columns.Count; j++)
                sHeaders = sHeaders.ToString() + Convert.ToString(d.Columns[j].HeaderText) + "\t";
            stOutput += sHeaders + "\r\n";
            // Export data.
            for (int i = 0; i < d.RowCount - 1; i++)
            {
                string stLine = "";
                for (int j = 0; j < d.Rows[i].Cells.Count; j++)
                    stLine = stLine.ToString() + Convert.ToString(d.Rows[i].Cells[j].Value) + "\t";
                stOutput += stLine + "\r\n";
            }
            Encoding utf16 = Encoding.GetEncoding(1254);
            byte[] output = utf16.GetBytes(stOutput);
            FileStream fs = new FileStream(filename, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(output, 0, output.Length); //write the encoded file
            bw.Flush();
            bw.Close();
            fs.Close();*/

            Ex.Application xlApp;
            Ex.Workbook xlWorkBook;
            Ex.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Ex.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Ex.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int i = 0;
            int j = 0;

            for (j = 0; j < d.ColumnCount; j++)
            {
                xlWorkSheet.Cells[i + 1, j + 1] = d.Columns[j].HeaderText;
                for (i = 0; i <= d.RowCount - 1; i++)
                {
                    DataGridViewCell cell = d[j, i];
                    xlWorkSheet.Cells[i + 2, j + 1] = cell.Value;
                }
                i = 0;
            }

            try
            {
                xlWorkBook.SaveAs(filename, Ex.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Ex.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            System.Diagnostics.Process.Start(filename);
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
    }
}