using Dev_Test.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Text.RegularExpressions;

namespace Dev_Test.Model
{
    class HammerSmith : Utils
    {
        private string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\tnyima\Downloads\Development Test\Sample\hammerSmithDB.accdb;Persist Security Info=False;";

        public void FlatFileReader(char delim, string fileLocation)
        {
            try
            {
                StreamReader streamReader = new StreamReader(fileLocation, System.Text.Encoding.Default);

                string contents = streamReader.ReadToEnd();
                string[] lines = System.Text.RegularExpressions.Regex.Split(contents, "\f", System.Text.RegularExpressions.RegexOptions.None);

                UpdateHeader();

                string accountNumber = "";
                int pageNumber = 0;
                int recordNumber = 0;
                bool firstpage = false;

                foreach (var line in lines)
                {
                    //pageNumber = 1;
 

                    var items = line.Split(new char[] { '\n' }, StringSplitOptions.None);

                    DataRow dr = myTable.NewRow();
                    int cnt = 1;

                    foreach (var item in items)
                    {
                        dr[$"Field{cnt}"] = item;

                        // Accounter Number
                        if (item.Contains(@"Account ref          "))
                        {
                            string[] accountText = Regex.Split(item, @"\s+");
                            accountNumber = accountText[7];
                        }

                        cnt++;
                    }

                    // Page Numbers

                    if(line.Contains("Date of issue")) 
                    {
                        firstpage = true;
                        pageNumber = 1;

                    } else pageNumber++;

                    if (line.Contains("LBHF_CTAXBILL$BILL")) {
                        recordNumber++;
                    } 

                    dr[$"AccNo"] = accountNumber;
                    dr[$"PageNo"] = pageNumber;
                    dr[$"Record"] = recordNumber;

                    myTable.Rows.Add(dr);
                }


                //    items.Add(record);
                // lines.Select(x => x.Split(new char[] { '\n' })).ToList().ForEach(row => myTable.Rows.Add(row));

                myTable.AsEnumerable().Select(x =>
                {
                    x["FileName"] = fileLocation;
                    return x;
                }).ToList();
                // Saving Datatable to Access Database#
                SaveDataToMSAccessDB();

            }
            catch (Exception ex) {
                MessageBox.Show("Error" + ex);
            }
        }

        public void UpdateHeader()
        {
            for(int col =1; col <= 100; col++)
            {
                myTable.Columns.Add($"Field{col}");
            }

            myTable.Columns.Add("FileName");
            myTable.Columns.Add("AccNo");
            myTable.Columns.Add("PageNo");
            myTable.Columns.Add("Record");
            myTable.Columns.Add("Profile");
        }

        // Saving Datatable to MS Access Database
        private void SaveDataToMSAccessDB() {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Clear table
                ClearAccessDatabase(connection);

                // Upload new data in table.
                var dataAdapter = new OleDbDataAdapter("SELECT * FROM import_table", connection);

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

        private static void ClearAccessDatabase(OleDbConnection connection)
        {
            OleDbCommand ac = new OleDbCommand("delete from import_table", connection);
            connection.Open();
            ac.ExecuteNonQuery();
            connection.Close();
        }
    }
}
