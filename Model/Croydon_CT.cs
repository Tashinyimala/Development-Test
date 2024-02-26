using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using Dev_Test.Util;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Dev_Test.Model
{
    class Croydon_CT: Utils
    {
        string[] headerLable = { "CLAIM", "FULL_NAME", "TITLE_SURNAME", "ADDRESS", "CLAIM_ADDRESS", "TODAY_DATE",
            "RENT_PRO_REFNO", "RENT_REQ_DATE", "RENT_DATE", "PIN_NO", "MESSAGEALL", "RENT_AMOUNT", "FREQ", "ADDRESS1",
            "ADDRESS2", "ADDRESS3","ADDRESS4","ADDRESS5","ADDRESS6","ADDRESS7" };

        private string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\tnyima\Downloads\Development Test\Sample\Croydon_CT.accdb;Persist Security Info=False;";

        string[] records;
        List<string> recordTexts = new List<string>();

        public void ReadFileData(char Delim)
        {
            if (File.Exists(FilePathname))
            {
                //string[] lines = System.IO.File.ReadAllLines(FilePathname);
                string lines = System.IO.File.ReadAllText(FilePathname);
                records = lines.Split(new string[] { "CLAIM			" }, StringSplitOptions.None);

                UpdateHeader();

                if (lines != null)
                {
                    ExtractDataUsingSeparator(Delim, records);
                }
            }

            SaveDataToMSAccessDB();
        }

        public virtual void ExtractDataUsingSeparator(char Delim, string[] records)
        {
            string[] dataWords = new List<String>().ToArray();
            List<string> recordTextTrimTab = new List<string>();
            int RecordNo = 0;

            for (int record = 1; record < records.Length; record++)
            {
                int PageNo = 0;
                
                recordTexts.Clear();
                recordTextTrimTab.Clear();

                dataWords = records[record]
                    .Split('\n')
                    .ToArray();

                recordTexts.Add(dataWords[0]); // Cliam number
                PageNo++;
                RecordNo++;

                foreach (string text in dataWords) {
                    recordTextTrimTab.Add(removeTabs(text));
                }

                foreach (string recordText in recordTextTrimTab)
                {
                    getFieldData(recordText, "FULL_NAME ");
                    getFieldData(recordText, "TITLE_SURNAME ");
                    getFieldData(recordText, "CLAIM_ADDRESS ");
                    getFieldData(recordText, "ADDRESS ");
                    getFieldData(recordText, "TODAY_DATE ");
                    getFieldData(recordText, "RENT_PRO_REFNO ");
                    getFieldData(recordText, "RENT_REQ_DATE ");
                    getFieldData(recordText, "RENT_DATE ");
                    getFieldData(recordText, "PIN_NO ");
                    getFieldData(recordText, "MESSAGEALL ");
                    getFieldData(recordText, "RENT_AMOUNT ");
                    getFieldData(recordText, "FREQ ");
                    getFieldData(recordText, "ADDRESS1 ");
                    getFieldData(recordText, "ADDRESS2 ");
                    getFieldData(recordText, "ADDRESS3 ");
                    getFieldData(recordText, "ADDRESS4 ");
                    getFieldData(recordText, "ADDRESS5 ");
                    getFieldData(recordText, "ADDRESS6 ");
                    getFieldData(recordText, "ADDRESS7 ");
                }

                recordTexts.RemoveAt(3); // Removing the duplidate address

                DataRow row = myTable.NewRow();

                int index = 0;
                while (index < recordTexts.Count) // loop through each line add the data to the row
                {
                    foreach (string data in recordTexts)
                    {
                        row[index] = data;
                        index++;
                    }
                }

                myTable.Rows.Add(row);

                myTable.AsEnumerable().Select(x =>
                {
                    x["MESSAGEALL"] = removeTabs(dataWords[11]);
                    return x;
                }).ToList();

                myTable.AsEnumerable().Select(x =>
                {
                    x["FileName"] = FilePathname;
                    return x;
                }).ToList();

                row[$"AccNo"] = dataWords[0];
                row[$"PageNo"] = PageNo;
                row[$"Record"] = RecordNo;
            }
        }

        private void getFieldData(string recordText, string searchText)
        {
            if (recordText.Contains(searchText))
            {
                string text = recordText;
                string toBeSearched = searchText;
                //string fieldToAdd = removeTabs(recordText.Substring(text.IndexOf(searchText) + toBeSearched.Length));
                string fieldToAdd = removeTabs(recordText.Replace(toBeSearched, string.Empty));


                recordTexts.Add(fieldToAdd);
            }
        }

        public string removeTabs(string strWithTabs) {
            char tab = '\u0009';
            String line = strWithTabs.Replace(tab.ToString(), " ");
            string result = Regex.Replace(line, @"\s+", " ");

            return result;
        }
        public void UpdateHeader()
        {
            for (int col = 0; col < headerLable.Length; col++)
            {
                myTable.Columns.Add(headerLable[col]);
            }

            myTable.Columns.Add("FileName");
            myTable.Columns.Add("AccNo");
            myTable.Columns.Add("PageNo");
            myTable.Columns.Add("Record");
        }

        // Saving Datatable to MS Access Database
        private void SaveDataToMSAccessDB()
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Clear table
                ClearAccessDatabase(connection);

                // Upload new data in table.
                var dataAdapter = new OleDbDataAdapter("SELECT * FROM Import_Table", connection);

                using (OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(dataAdapter))
                {
                    try
                    {
                        connection.Open();
                        commandBuilder.ConflictOption = ConflictOption.CompareRowVersion;
                        dataAdapter.Update(myTable);
                    }
                    catch (OleDbException ex)
                    {
                        MessageBox.Show(ex.Message, "OleDbException Error");
                    }
                    catch (Exception x)
                    {
                        MessageBox.Show(x.Message, "Exception Error");
                    }
                }

                connection.Close();
                dataAdapter.Dispose();
            }
        }

        public void updateDB(DataTable xDataTable)
        {
            var dataAdapter = new OleDbDataAdapter("SELECT * FROM Import_Table", connectionString);


            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Clear table
                //ClearAccessDatabase(connection);

                using (OleDbCommandBuilder OleCB = new OleDbCommandBuilder(dataAdapter))
                {
                    connection.Open();
                    OleCB.ConflictOption = ConflictOption.CompareRowVersion;
                    dataAdapter.Update(xDataTable);
                }

                // Get Updated records
                //myTable.Clear();
                //dataAdapter.Fill(myTable);

                connection.Close();
                dataAdapter.Dispose();

                MessageBox.Show("Data Updated successfully.....");
            }
        }

        private static void ClearAccessDatabase(OleDbConnection connection)
        {
            OleDbCommand ac = new OleDbCommand("delete from Import_Table", connection);
            connection.Open();
            ac.ExecuteNonQuery();
            connection.Close();
        }
    }
}
