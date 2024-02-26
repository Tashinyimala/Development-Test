using Dev_Test.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Dev_Test.Model
{
    class Books : Utils
    {
        //OleDBConnection for MS Access Database
        private OleDbConnection connection = new OleDbConnection();

        public void XMLFileReader(string fileLocation)
        {
            XmlDocument xmlDocument = new XmlDocument();

            try
            {
                xmlDocument.Load(fileLocation);
                XmlNodeList nodes = xmlDocument.DocumentElement.SelectNodes("/catalog/book");
               
                string id = "", author = "", title = "", genre = "", price = "", publish_date = "", description = "";

                foreach (XmlNode node in nodes)
                {
                    id = node.Attributes["id"].Value;
                    author = node.SelectSingleNode("author").InnerText;
                    title = node.SelectSingleNode("title").InnerText;
                    genre = node.SelectSingleNode("genre").InnerText;
                    price = node.SelectSingleNode("price").InnerText;
                    publish_date = node.SelectSingleNode("publish_date").InnerText;
                    description = node.SelectSingleNode("description").InnerText;

                    string[] dataWords = { id, author, title, genre, price, publish_date, description };

                    SaveDataToMSAccessDB(dataWords);
                }
            }
            finally
            {
                connection.Close();
            }
        }

        // Saving data to access Database
        private void SaveDataToMSAccessDB(string[] dataWords)
        {
            OleDbCommand command = ConnectToAccessDatabase();

            // Command with placeholders
            command.CommandText = "INSERT INTO BooksDB" + "([ID], [Author], [Title], [Genre], [Price], [Publish Date], [Description])" +
                "VALUES(@iD, @author, @title, @genre, @price, @publish_date, @description)";

            // add named parameters
            command.Parameters.AddRange(new OleDbParameter[] 
            {
                            new OleDbParameter("@iD", dataWords[0]),
                            new OleDbParameter("@author", dataWords[1]),
                            new OleDbParameter("@title", dataWords[2]),
                            new OleDbParameter("@genre", dataWords[3]),
                            new OleDbParameter("@price", dataWords[4]),
                            new OleDbParameter("@publish_date", dataWords[5]),
                            new OleDbParameter("@description", dataWords[6]),
            });

            command.ExecuteNonQuery();

            connection.Close();
        }

        // MS Access Database connection
        private OleDbCommand ConnectToAccessDatabase()
        {
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\tnyima\Downloads\Development Test\Sample\XML_Books_DB.accdb;Persist Security Info=False;";
            connection.Open();

            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            return command;
        }

        // Reading all records from Access Database
        public void ReadDataFromAccessDatabase()
        {
            try
            {
                OleDbCommand command = ConnectToAccessDatabase();
                command.CommandText = "SELECT * From BooksDB";

                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command);
                dataAdapter.Fill(myTable);

                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex);
            }

        }
    }
}
