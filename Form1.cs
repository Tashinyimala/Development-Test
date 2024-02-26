using Dev_Test.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Dev_Test
{
    public partial class Form1 : Form
    {
        string filePathName = "";
        string fileName = "";
        char charDeliminator;

        Employee employee = new Employee();
        Electoral electoral = new Electoral();
        Client_xlsx client = new Client_xlsx();
        FlatFile_HB flatFile = new FlatFile_HB();
        Books xmlBooks = new Books();
        Enfield xmlEnfield = new Enfield();
        Croydon croydon = new Croydon();
        HammerSmith hammerSmith = new HammerSmith();
        Luton luton = new Luton();
        Enfield_CT enfield_CT = new Enfield_CT();
        Croydon_CT croydon_CT = new Croydon_CT();

        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            DialogResult result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                filePathName = openFileDialog.FileName; // Get full file path
                fileName = System.IO.Path.GetFileName(filePathName);

                try
                {
                    lblFilePath.Text = filePathName;
                    lblFilePath.Visible = true;
                    employee.FilePathname = filePathName;
                    electoral.FilePathname = filePathName;
                    croydon_CT.FilePathname = filePathName;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("File can't be found");
                    Console.WriteLine(ex.Message);
                }
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            string deliminator = charDeliminator.ToString();
            Cursor = Cursors.WaitCursor;
            //if (deliminator != "\0") // "\0" is for blank. If not blank then if statement are executed.
            //{
                switch (fileName)
                {
                    case "01 data_comma.csv":
                        employee.ReadCommaSepratedFle(charDeliminator);
                        updateEmployeeTable();
                        break;

                    case "02 data_comma_quoted.csv":
                        electoral.ClearTable();
                        electoral.ReadFile(charDeliminator);
                        myDataGridView.DataSource = electoral.myTable;
                        myDataGridViewPropertySetting();
                        break;

                    case "03 data_tab.txt":
                        electoral.ClearTable();
                        electoral.ReadFile(charDeliminator);
                        myDataGridView.DataSource = electoral.myTable;
                        myDataGridViewPropertySetting();
                        break;

                    case "04 Sample.xlsx":
                        client.ClearTable();
                        client.ExcelReader(charDeliminator, filePathName);
                        client.ReadDataFromAccessDatabase();
                        myDataGridView.DataSource = client.myTable;
                        myDataGridViewPropertySetting();
                        break;

                    case "05 data_flatfile_hb.dat":
                        flatFile.ClearTable();
                        flatFile.FlatFileReader(charDeliminator, filePathName);
                        myDataGridView.DataSource = flatFile.myTable;
                        break;

                    case "08 data_xml.xml":
                        xmlBooks.ClearTable();
                        xmlBooks.XMLFileReader(filePathName);
                        xmlBooks.ReadDataFromAccessDatabase();
                        myDataGridView.DataSource = xmlBooks.myTable;
                        myDataGridViewPropertySetting();
                        break;

                    case "11 data_xml_complex.xml":
                        xmlEnfield.ClearTable();
                        xmlEnfield.XMLFileReader(filePathName);
                        myDataGridView.DataSource = xmlEnfield.myTable;
                        break;

                    case "06 croyndrrems_5516575.txt":
                        croydon.ClearTable();
                        croydon.FlatFileReader(charDeliminator,filePathName);
                        myDataGridView.DataSource = croydon.myTable;
                        break;

                    case "06 data_flatfile_ct":
                        hammerSmith.ClearTable();
                        hammerSmith.FlatFileReader(charDeliminator, filePathName);
                        myDataGridView.DataSource = hammerSmith.myTable;
                        break;

                    case "07 data_PDF.pdf":
                        luton.ClearTable();
                        luton.FlatFileReader(charDeliminator, filePathName);
                        myDataGridView.DataSource = luton.myTable;
                        break;

                    case "09 data_RTF.rtf":
                        enfield_CT.ClearTable();
                        enfield_CT.FlatFileReader(charDeliminator, filePathName);
                        myDataGridView.DataSource = enfield_CT.myTable;
                        break;

                    case "12 lbc_roa_ltr_4704460.txt":
                        croydon_CT.ClearTable();
                        croydon_CT.ReadFileData(charDeliminator);
                        myDataGridView.DataSource = croydon_CT.myTable;
                        myDataGridViewPropertySetting();
                    break;

                default:
                        MessageBox.Show("Select correct file");
                        break;
                }
            //}
            //else
            //{
            //    MessageBox.Show("Please enter correct delimiter , tab,");
            //}
            Cursor = Cursors.Arrow;
        }

        private void updateEmployeeTable()
        {
            
            myDataGridView.ReadOnly = true;

            // Resize column based on content of the cell.
            myDataGridView.AutoResizeColumns();
            myDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            ExtractEmployeeData();
        }

        private void ExtractEmployeeData()
        {
            // Getting data from the file through employee model
            if (employee.employeeList.Count > 0)
            {
                myDataGridView.DataSource = employee.employeeList;
                UpdateHeaderText(); // Update the header
                myDataGridView.Columns["FilePathname"].Visible = false;
                myDataGridViewPropertySetting();

            }
            else
            {
                MessageBox.Show("Incorrect file.");
            }
        }

        // Set DataGridView Perperties
        private void myDataGridViewPropertySetting()
        {
            myDataGridView.ReadOnly = false;

            // Resize column based on content of the cell.
            myDataGridView.AutoResizeColumns();
            myDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        /// <summary>
        /// Update the DataGridView Header - Custom Header Text
        /// </summary>
        private void UpdateHeaderText()
        {
            string[] headerLables = employee.headerLables;

            int index = 0;

            foreach (string label in headerLables)
            {
                myDataGridView.Columns[index].HeaderText = headerLables[index];
                index++;
            }
        }

        /// <summary>
        /// Get the user entered delimiter and assign to charDelinater
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //private void txtDelimi_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    charDeliminator = e.KeyChar;

        //    // Dealing with Tab character
        //    if (e.KeyChar == '\t' || e.KeyChar == (char)13)
        //        e.Handled = true;
        //}

        /// <summary>
        /// Check user enter character is comma, if it is then import button is shown 
        /// Otherwise hide it.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //private void txtDelimi_KeyUp(object sender, KeyEventArgs e)
        //{
        //    //Console.WriteLine(e.KeyCode.ToString());
        //    //if (Regex.IsMatch(e.KeyCode.ToString(), @"Oemcomma"))
        //    //{
        //    //    btnImport.Visible = true;
        //    //    e.Handled = true;
        //    //}
        //    //else
        //    //{
        //    //    MessageBox.Show("Please enter correct delimiter , tab, ");
        //    //    btnImport.Visible = false;
        //    //    txtDelimi.Text = String.Empty;
        //    //    txtDelimi.Refresh();
        //    //}
        //}

        //private void txtDelimi_KeyDown(object sender, KeyEventArgs e)
        //{
        //    //if (e.KeyCode == Keys.Tab)
        //    //{
        //    //    txtDelimi.AppendText(@"\t");
        //    //}
        //}

        //private void txtDelimi_TextChanged(object sender, EventArgs e)
        //{
        //    //txtDelimi.Multiline = true;
        //    //txtDelimi.AcceptsReturn = true;
        //    //txtDelimi.AcceptsTab = true;
        //}

        private void cboxDelimiter_TextChanged(object sender, EventArgs e)
        {
            string selected = cboxDelimiter.GetItemText(cboxDelimiter.SelectedItem);

            if (selected == "No Delimiter")
            {
                //cboxDelimiter.Enabled = false;
            }
            else {
                charDeliminator = selected == "tab" ? '\t' : char.Parse(selected);
            }
        }



        private void btnDBUpdate_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = (DataTable)myDataGridView.DataSource;

            croydon_CT.updateDB(dt);
        }
    }
}
