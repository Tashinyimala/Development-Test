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
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace Dev_Test.Model
{
    class Enfield_CT: Utils
    {
        private string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\tnyima\Downloads\Development Test\Sample\Enfield_CT.accdb;Persist Security Info=False;";

        public void FlatFileReader(char delim, string fileLocation)
        {
            try
            {
                string accountNumber = "";
                int pageNumber = 0;
                int recordNumber = 0;

                string pdfFileName = @"C:\Users\tnyima\Downloads\Development Test\Sample\09 data_RTF.pdf";

                UpdateHeader();

                // Converting rtf to pdf
                //ConvertRTFtoPDF(fileLocation);

                PdfReader reader = new PdfReader(pdfFileName);
                int intPageNum = reader.NumberOfPages;
                string[] words;
                string line;

                for (int i = 1; i <= intPageNum; i++)
                {
                    string text = PdfTextExtractor.GetTextFromPage(reader, i, new LocationTextExtractionStrategy());
                    words = text.Split('\n');

                    DataRow dr = myTable.NewRow();
                    int cnt = 1;

                    dr[$"Field{cnt}"] = text;

                    foreach (var item in words)
                    {
                        line = Encoding.UTF8.GetString(Encoding.UTF8.GetBytes(item));

                        if (line.Contains(@"Claim Reference:   "))
                        {
                            pageNumber = 0;
                            recordNumber++;
                        }
                        else {

                        }

                        if (line.Contains(@"Council Tax Account:"))
                        {
                            accountNumber = line.Substring(line.Length - 9);
                        }

                        //// Accounter Number
                        //if (i % 2 == 0)
                        //{
                        //    string currentAcc = accountNumber;
                        //    accountNumber = currentAcc;
                        //}
                        //else
                        //{

                        //}


                    }


                    pageNumber++;
                    cnt++;
                    dr[$"AccNo"] = accountNumber;
                    dr[$"PageNo"] = pageNumber;
                    dr[$"Record"] = recordNumber;

                    myTable.Rows.Add(dr);
                }

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


            //new FileInfo("09 data_RTF.pdf").Delete();
        }

        private static void ConvertRTFtoPDF(string fileLocation)
        {
            Word.Application wordApp = new Word.Application();
            object missing = System.Reflection.Missing.Value;

            FileInfo rtfFile = new FileInfo(fileLocation);

            wordApp.Visible = false;
            wordApp.ScreenUpdating = false;

            object fileName = (object)rtfFile.FullName;


            Word.Document wordDoc = wordApp.Documents.Open(ref fileName, ref missing, 
                                                            ref missing, ref missing, ref missing, ref missing, ref missing,
                                                            ref missing, ref missing, ref missing, ref missing, ref missing,
                                                            ref missing, ref missing, ref missing, ref missing);
            wordDoc.Activate();

            object outputFilename = rtfFile.FullName.Replace(".rtf", ".pdf");
            object fileFormat = WdSaveFormat.wdFormatPDF;
            wordDoc.SaveAs2(ref outputFilename,
                            ref fileFormat, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing);

            object savechanges = WdSaveOptions.wdSaveChanges;

            wordDoc.Close(true);
            wordApp.Quit();

            MessageBox.Show("RTF file converted.");
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
        }

        // Saving Datatable to MS Access Database
        private void SaveDataToMSAccessDB() {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Clear table
                ClearAccessDatabase(connection);

                // Upload new data in table.
                var dataAdapter = new OleDbDataAdapter("SELECT * FROM import_data", connection);

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
            OleDbCommand ac = new OleDbCommand("delete from import_data", connection);
            connection.Open();
            ac.ExecuteNonQuery();
            connection.Close();
        }
    }
}
