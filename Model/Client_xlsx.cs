using Dev_Test.Util;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Dev_Test.Model
{
    class Client_xlsx : Utils
    {
        //OleDBConnection for MS Access Database
        private OleDbConnection connection = new OleDbConnection();

        public void ExcelReader(char delim, string fileLocation)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWB = excelApp.Workbooks.Open(fileLocation);
            excelWB.SaveAs(fileLocation + ".csv", Excel.XlFileFormat.xlCSVWindows);
            excelWB.Close(true);

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWB);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);

            List<string> valueList = null;
            using (StreamReader streamReader = new StreamReader(fileLocation + ".csv"))
            {
                string content = streamReader.ReadToEnd();
                valueList = new List<string>(
                    content.Split(
                        new string[] { "\r\n" },
                        StringSplitOptions.RemoveEmptyEntries
                    )
                );

                // Extracting each cell's data, seprated by commna
                for (int line = 1; line < valueList.Count; line++)
                {
                    string[] dataWords = valueList[line].Split(delim);

                    // Adding data to MS Access Database
                    try
                    {
                        SaveDataToMSAccessDB(dataWords);
                    }
                    catch (Exception ex) {
                        MessageBox.Show("Error"+ ex);
                    }
                }
            }

            new FileInfo(fileLocation + ".csv").Delete();
            MessageBox.Show("Updated all the data to access database");
        }


        // MS Access Database connection
        private OleDbCommand ConnectToAccessDatabase()
        {
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\tnyima\Downloads\Development Test\Sample\DevTest_DB.accdb;Persist Security Info=False;";
            connection.Open();

            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            return command;
        }

        // Saving data to access Database
        private void SaveDataToMSAccessDB(string[] dataWords)
        {
            OleDbCommand command = ConnectToAccessDatabase();

            // Command with placeholders
            command.CommandText = "INSERT INTO DevTestTable" + "([ID], [Field1], [Field2], [Field3], [Field4], [Field5], [Field6], [Field7]) " +
                "VALUES(@iD, @field1, @field2, @field3, @field4, @field5, @field6, @field7)";

            // add named parameters
            command.Parameters.AddRange(new OleDbParameter[] {
                            new OleDbParameter("@iD", dataWords[0]),
                            new OleDbParameter("@field1", dataWords[1]),
                            new OleDbParameter("@field2", dataWords[2]),
                            new OleDbParameter("@field3", dataWords[3]),
                            new OleDbParameter("@field4", dataWords[4]),
                            new OleDbParameter("@field5", dataWords[5]),
                            new OleDbParameter("@field6", dataWords[6]),
                            new OleDbParameter("@field7", dataWords[7]),
             });

            command.ExecuteNonQuery();

            connection.Close();
        }


        // Reading all records from Access Database
        public void ReadDataFromAccessDatabase()
        {
            try
            {
                OleDbCommand command = ConnectToAccessDatabase();
                command.CommandText = "SELECT * From DevTestTable";

                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command);
                dataAdapter.Fill(myTable);

                connection.Close();
            } catch (Exception ex)
            {
                MessageBox.Show("Error" + ex);
            }

        }
    }
}
