using Dev_Test.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Dev_Test.Model
{
    class Croydon: Utils
    {
        private string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\tnyima\Downloads\Development Test\Sample\Croydon.accdb;Persist Security Info=False;";

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

                foreach (var line in lines)
                {
                    pageNumber = 0;

                    var items = line.Split(new char[] { '\n' }, StringSplitOptions.None);

                    DataRow datarow = myTable.NewRow();
                    int cnt = 1;
                    foreach (var item in items)
                    {
                        datarow[$"field{cnt}"] = item;
                        cnt++;
                    }

                    // Extracing Account Numbers
                    if (line.Contains(@"REMINDER"))
                    {
                        accountNumber = line.Substring(21, 8);
                    }

                    //accountNumber = items[2];
                    pageNumber++;
                    recordNumber++;
                
                    datarow[$"AccNo"] = accountNumber;
                    datarow[$"PageNo"] = pageNumber;
                    datarow[$"Records"] = recordNumber;

                    myTable.Rows.Add(datarow);
                }

                // lines.Select(x => x.Split(new char[] { '\n' })).ToList().ForEach(row => myTable.Rows.Add(row));

                myTable.AsEnumerable().Select(x => {
                    x["FileName"] = fileLocation;
                    return x;
                }).ToList();


                // Saving Datatable to Access Database#
                SaveDataToMSAccessDB();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex);
            }
        }


        private void UpdateHeader()
        {
            for (int col = 1; col <= 100; col++)
            {
                myTable.Columns.Add($"Field{col}");
            }

            myTable.Columns.Add("FileName");
            myTable.Columns.Add("AccNo");
            myTable.Columns.Add("PageNo");
            myTable.Columns.Add("Records");
            myTable.Columns.Add("Profile");

        }

        private void SaveDataToMSAccessDB()
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {

                // Clear table
                ClearAccessDatabase(connection);

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

        private static void ClearAccessDatabase(OleDbConnection connection)
        {
            OleDbCommand accessDBconnection = new OleDbCommand("delete from Import_Table", connection);
            connection.Open();
            accessDBconnection.ExecuteNonQuery();
            connection.Close();
        }

    }
}
