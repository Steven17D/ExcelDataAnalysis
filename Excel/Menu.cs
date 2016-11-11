using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public partial class Menu : Form
    {
        private OpenFileDialog ofd = new OpenFileDialog();
        private string dataType;
        private OleDbConnection connection;
        private string con;

        public Menu()
        {
            InitializeComponent();
        }

        private void openFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                pathLable.Text = ofd.FileName;
                sheetSelect.Items.Clear();
                getData();
            }
        }

        private void getData()
        {
            //Specify the excel provider for .xlsx file type and the file path which contain the excel file
            con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathLable.Text + ";Extended Properties=Excel 12.0 Xml";
            connection = new OleDbConnection(con);
            using (connection)
            {
                //Open the Oledb connection
                connection.Open();
                DataTable dtSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                foreach (DataRow table in dtSchema.Rows)
                {
                    sheetSelect.Items.Add(table.Field<string>("TABLE_NAME"));
                }
                sheetSelect.Visible = true;
                connection.Close();
            }
        }

        private void Menu_Resize(object sender, EventArgs e)
        {
            if (this.Size.Width - 40 < dataGridView.PreferredSize.Width - 32)
            {
                dataGridView.Size = new Size(this.Size.Width - 40, this.Size.Height - 96);
            }
            if (this.Size.Width - 40 >= dataGridView.PreferredSize.Width - 32)
            {
                dataGridView.Size = new Size(dataGridView.PreferredSize.Width - 32, this.Size.Height - 96);
            }
            if (this.Size.Height - 96 >= dataGridView.PreferredSize.Height - 32)
            {
                dataGridView.Size = new Size(this.Size.Width - 40, dataGridView.PreferredSize.Height - 32);
            }
            if (this.Size.Height - 96 >= dataGridView.PreferredSize.Height - 32 && this.Size.Width - 40 >= dataGridView.PreferredSize.Width - 32)
            {
                dataGridView.Size = new Size(dataGridView.PreferredSize.Width - 32, dataGridView.PreferredSize.Height - 32);
            }
        }

        private void sheetSelect_SelectedValueChanged(object sender, EventArgs e)
        {
            calculateSummeryToolStripMenuItem.Enabled = true;
            using (connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [" + sheetSelect.SelectedItem + "]", connection);
                //Define the object of data adaper to run the query
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                //Define the dataset to hold the records
                DataSet ds = new DataSet();
                //Fill the data set
                adap.Fill(ds);
                //Check the condition if dataset contains any table, It should be at least greater than one
                if (ds.Tables.Count >= 1)
                {
                    dataGridView.DataSource = ds.Tables[0];
                    //dataGridView.DataBind();
                }
                connection.Close();
                this.Size = new Size(500, 500);
                this.MaximumSize = new Size(dataGridView.PreferredSize.Width + 15, dataGridView.PreferredSize.Height + 72);
            }
        }

        private void sheetSelect_MouseWheel(object sender, MouseEventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }

        private void calculateSummeryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            calculateSummeryToolStripMenuItem.Enabled = false;
            if (dataGridView.ColumnCount + dataGridView.RowCount == 0)
            {
                MessageBox.Show("ERROR: There is no data!");
                return;
            }

            #region Clean the sheet from unnecessary information

            if (dataGridView.Columns.Contains("Total_Body_Scrap"))
            {
                dataType = "Body";
                try
                {
                    if (dataGridView.Columns.Contains("ID"))
                        dataGridView.Columns.Remove("ID"); //id
                    if (dataGridView.Columns.Contains("StartTime_Run"))
                        dataGridView.Columns.Remove("StartTime_Run"); //data
                    if (dataGridView.Columns.Contains("Shift"))
                        dataGridView.Columns.Remove("Shift"); //shift
                    //if (dataGridView.Columns.Contains("RunName"))
                    //    dataGridView.Columns.Remove("RunName");
                    if (dataGridView.Columns.Contains("Total_Body_Scrap"))
                        dataGridView.Columns.Remove("Total_Body_Scrap"); //for BODY sheet
                    if (dataGridView.Columns.Contains("BodyID_1"))
                        dataGridView.Columns.Remove("BodyID_1");
                    if (dataGridView.Columns.Contains("BodyID_2"))
                        dataGridView.Columns.Remove("BodyID_2");
                    if (dataGridView.Columns.Contains("BodyID_3"))
                        dataGridView.Columns.Remove("BodyID_3");
                    if (dataGridView.Columns.Contains("BodyID_4"))
                        dataGridView.Columns.Remove("BodyID_4");
                    if (dataGridView.Columns.Contains("BodyID_5"))
                        dataGridView.Columns.Remove("BodyID_5");
                    if (dataGridView.Columns.Contains("BodyID_6"))
                        dataGridView.Columns.Remove("BodyID_6");
                    if (dataGridView.Columns.Contains("Oprator"))
                        dataGridView.Columns.Remove("Oprator");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    Application.Restart();
                    return;
                }
            }
            else
            {
                dataType = "CSL";
                try
                {
                    if (dataGridView.Columns.Contains("ID"))
                        dataGridView.Columns.Remove("ID"); //id
                    if (dataGridView.Columns.Contains("StartTime_Run"))
                        dataGridView.Columns.Remove("StartTime_Run"); //data
                    if (dataGridView.Columns.Contains("Shift"))
                        dataGridView.Columns.Remove("Shift"); //shift
                    //if (dataGridView.Columns.Contains("RunName"))
                    //    dataGridView.Columns.Remove("RunName");
                    if (dataGridView.Columns.Contains("Total_CSL_Scrap"))
                        dataGridView.Columns.Remove("Total_CSL_Scrap");
                    if (dataGridView.Columns.Contains("CSLID_1"))
                        dataGridView.Columns.Remove("CSLID_1");
                    if (dataGridView.Columns.Contains("CSLID_2"))
                        dataGridView.Columns.Remove("CSLID_2");
                    if (dataGridView.Columns.Contains("CSLID_3"))
                        dataGridView.Columns.Remove("CSLID_3");
                    if (dataGridView.Columns.Contains("CSLID_4"))
                        dataGridView.Columns.Remove("CSLID_4");
                    if (dataGridView.Columns.Contains("CSLID_5"))
                        dataGridView.Columns.Remove("CSLID_5");
                    if (dataGridView.Columns.Contains("CSLID_6"))
                        dataGridView.Columns.Remove("CSLID_6");
                    if (dataGridView.Columns.Contains("Oprator"))
                        dataGridView.Columns.Remove("Oprator");
                }
                catch (Exception exe)
                {
                    MessageBox.Show(exe.Message);
                    Application.Restart();
                    return;
                }
            }

            //iterate over the ProductName
            for (int rows = 0; rows < dataGridView.Rows.Count; rows++)
            {
                dataGridView.Rows[rows].Cells[0].ToolTipText = rows.ToString();
                string temp = dataGridView.Rows[rows].Cells[0].Value.ToString();
                switch (temp)
                {
                    case "Gemini CS - 300":
                        dataGridView.Rows[rows].Cells[0].Value = "Gemini CS";
                        break;

                    case "Gemini  CS - 300":
                        dataGridView.Rows[rows].Cells[0].Value = "Gemini CS";
                        break;

                    case "Gemini CS - 320":
                        dataGridView.Rows[rows].Cells[0].Value = "Gemini CS";
                        break;

                    case "Gemini  CS - 320":
                        dataGridView.Rows[rows].Cells[0].Value = "Gemini CS";
                        break;

                    case "Gemini Long - 360":
                        dataGridView.Rows[rows].Cells[0].Value = "Gemini";
                        break;

                    case "Gemini  Long - 360":
                        dataGridView.Rows[rows].Cells[0].Value = "Gemini";
                        break;

                    case "Gemini Short - 180":
                        dataGridView.Rows[rows].Cells[0].Value = "Gemini";
                        break;

                    case "Gemini  Short - 180":
                        dataGridView.Rows[rows].Cells[0].Value = "Gemini";
                        break;

                    case "Iris plus CS - 320":
                        dataGridView.Rows[rows].Cells[0].Value = "Iris plus CS";
                        break;

                    case "Iris plus Long - 360":
                        dataGridView.Rows[rows].Cells[0].Value = "Iris plus";
                        break;

                    case "Iris plus Short - 180":
                        dataGridView.Rows[rows].Cells[0].Value = "Iris plus";
                        break;

                    case "Iris plus  Short - 180":
                        dataGridView.Rows[rows].Cells[0].Value = "Iris plus";
                        break;

                    case "Iris std Long - 360":
                        dataGridView.Rows[rows].Cells[0].Value = "Iris std";
                        break;

                    case "Iris std Short - 180":
                        dataGridView.Rows[rows].Cells[0].Value = "Iris std";
                        break;

                    case "Iris std  Short - 180":
                        dataGridView.Rows[rows].Cells[0].Value = "Iris std";
                        break;

                    case "Iris Long - 360":
                        if (dataGridView.Rows[rows].Cells[1].Value.ToString().Length < 2)
                        {
                            dataGridView.Rows.Remove(dataGridView.Rows[rows]);
                            rows--;
                        }
                        if (dataGridView.Rows[rows].Cells[1].Value.ToString()[1] == 'S')
                        {
                            dataGridView.Rows[rows].Cells[0].Value = "Iris std";
                        }
                        else if (dataGridView.Rows[rows].Cells[1].Value.ToString()[1] == 'D')
                        {
                            dataGridView.Rows[rows].Cells[0].Value = "Iris plus";
                        }
                        else
                        {
                            //DialogResult dr = MessageBox.Show("Bad run name for product: " + dataGridView.Rows[rows].Cells[0].Value.ToString()
                            //    + "\nRun name: " + dataGridView.Rows[rows].Cells[1].Value.ToString() + "\nRow was deleted!"
                            //, "Bad run name", MessageBoxButtons.OK);
                            //dataGridView.Rows[rows].Cells[0].Style.BackColor = Color.Red;
                            //dataGridView.Rows.Remove(dataGridView.Rows[rows]);
                            //rows--; //after deletion the amount of rows decrises by one so, we need to loop one back
                            dataGridView.Rows[rows].Cells[0].Value = "Iris unknown";
                        }
                        break;

                    case "Iris Short - 180":
                        if (dataGridView.Rows[rows].Cells[1].Value.ToString().Length < 2)
                        {
                            dataGridView.Rows.Remove(dataGridView.Rows[rows]);
                            rows--;
                        }
                        if (dataGridView.Rows[rows].Cells[1].Value.ToString()[1] == 'S')
                        {
                            dataGridView.Rows[rows].Cells[0].Value = "Iris std";
                        }
                        else if (dataGridView.Rows[rows].Cells[1].Value.ToString()[1] == 'D')
                        {
                            dataGridView.Rows[rows].Cells[0].Value = "Iris plus";
                        }
                        else
                        {
                            //DialogResult dr = MessageBox.Show("Bad run name for product: " + dataGridView.Rows[rows].Cells[0].Value.ToString()
                            //    + "\nRun name: " + dataGridView.Rows[rows].Cells[1].Value.ToString() + "\nRow was deleted!"
                            //, "Bad run name", MessageBoxButtons.OK);
                            //dataGridView.Rows[rows].Cells[0].Style.BackColor = Color.Red;
                            //dataGridView.Rows.Remove(dataGridView.Rows[rows]);
                            //rows--; //after deletion the amount of rows decrises by one so, we need to loop one back
                            dataGridView.Rows[rows].Cells[0].Value = "Iris unknown";
                        }
                        break;

                    case "K2X - 180":
                        dataGridView.Rows[rows].Cells[0].Value = "K2X";
                        break;

                    case "K23 - 180":
                        dataGridView.Rows[rows].Cells[0].Value = "K2X";
                        break;

                    case "K2Y - 180":
                        dataGridView.Rows[rows].Cells[0].Value = "K2X";
                        break;

                    case "K2X - 000":
                        dataGridView.Rows[rows].Cells[0].Value = "K2X";
                        break;

                    case "O3 - 180":
                        dataGridView.Rows[rows].Cells[0].Value = "O3";
                        break;

                    case "Rotem Long - 220":
                        dataGridView.Rows[rows].Cells[0].Value = "Rotem";
                        break;

                    case "Rotem Short - 110":
                        dataGridView.Rows[rows].Cells[0].Value = "Rotem";
                        break;

                    case "Timna Long - 220":
                        dataGridView.Rows[rows].Cells[0].Value = "Timna";
                        break;

                    case "Timna Short - 110":
                        dataGridView.Rows[rows].Cells[0].Value = "Timna";
                        break;

                    case "Timna 2 Long - 220":
                        dataGridView.Rows[rows].Cells[0].Value = "Timna 2";
                        break;

                    case "Timna 2 Short - 000":
                        dataGridView.Rows[rows].Cells[0].Value = "Timna 2";
                        break;

                    case "Gemini 3 Long - 360":
                        dataGridView.Rows[rows].Cells[0].Value = "Gemini 3";
                        break;

                    case "Gemini 3 Short - 000":
                        dataGridView.Rows[rows].Cells[0].Value = "Gemini 3";
                        break;

                    case "Iris plus CS - 300":
                        dataGridView.Rows[rows].Cells[0].Value = "Iris plus CS";
                        break;

                    case "Iris CS - 320":
                        dataGridView.Rows[rows].Cells[0].Value = "Iris CS";
                        break;

                    case "":
                        dataGridView.Rows.Remove(dataGridView.Rows[rows]);
                        rows--;
                        break;

                    default:
                        DialogResult dialogResult = MessageBox.Show("Got bad product name: " + dataGridView.Rows[rows].Cells[0].Value.ToString() + "\nDelete the product row?"
                            , "Bad Product", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            dataGridView.Rows.Remove(dataGridView.Rows[rows]);
                            rows--; //after deletion the amount of rows decrises by one so, we need to loop one back
                        }
                        break;
                }
            }
            //return;
            if (dataGridView.Columns.Contains("RunName"))
                dataGridView.Columns.Remove("RunName");

            #endregion Clean the sheet from unnecessary information

            #region Sums all the scraps

            /*Dictionary:
             *          name: ((resone, length), (resone, length), ...) ,
             *          name: ...
            */
            Dictionary<string, Dictionary<string, double>> mainInformation = new Dictionary<string, Dictionary<string, double>>();
            string productName, scrapName;
            double scrapValue;
            Dictionary<string, double> productInfo;
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                productName = (string)row.Cells[0].Value;
                if (!mainInformation.ContainsKey(productName))
                {
                    mainInformation.Add(productName, new Dictionary<string, double>());
                }
                productInfo = mainInformation[productName];
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    try
                    {
                        DataGridViewCell cell = row.Cells[i];
                        if (cell.Size.IsEmpty) { continue; }
                        if (cell.ValueType == typeof(double))
                        {
                            if ((double)cell.Value != 0)
                            {
                                try
                                {
                                    scrapName = (string)row.Cells[i + 1].Value;
                                }
                                catch (Exception)
                                {
                                    scrapName = "No reason";
                                }
                                scrapValue = (double)cell.Value;
                                if (productInfo.ContainsKey(scrapName))
                                {
                                    productInfo[scrapName] += scrapValue;
                                }
                                else
                                {
                                    productInfo.Add(scrapName, scrapValue);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message);
                    }
                }
            }

            #endregion Sums all the scraps

            #region Display all the info in a DataGridView

            Summary dataDisplay = new Summary(mainInformation, dataType);
            dataDisplay.Show();

            #endregion Display all the info in a DataGridView
        }
    }
}