using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dev_Test.Util
{
    class Utils
    {
        private string _FilePathname;

        private DataTable _myTable = new DataTable();

        public string[] headerLables = new List<String>().ToArray();

        public DataTable myTable
        {
            get => _myTable;

            set => _myTable = value;
        }

        public string FilePathname
        {
            get => _FilePathname;
            set => _FilePathname = value;
        }



        public void ReadFile(char Delim)
        {
            if (File.Exists(FilePathname))
            {
                string[] lines = System.IO.File.ReadAllLines(FilePathname);
                if (lines.Length > 0)
                {
                    ExtractHeader(Delim, lines);
                    ExtractDataExcludingHeader(Delim, lines);
                }
            }
        }

        public void ExtractHeader(char Delim, string[] lines)
        {
            // Extracting Header from the file
            string firstLine = lines[0];
            //headerLables = firstLine.Split(Delim);
            headerLables = firstLine
                            .Split(Delim)
                            .Select(p => p.Trim(Delim))
                            .Where(p => !string.IsNullOrWhiteSpace(p))
                            .ToArray();

            foreach (string header in headerLables)
            {
                myTable.Columns.Add(header);
            }
        }

        public virtual void ExtractDataExcludingHeader(char Delim, string[] lines) { }

        public void ClearTable()
        {
            if (myTable != null && myTable.Rows.Count > 0 )
            {
                myTable.Clear();
            }
        }

        public virtual void ExtractDataUsingSeparator(char Delim, string[] lines, char separator)
        {
            string[] dataWords = new List<String>().ToArray();

            for (int line = 1; line < lines.Length; line++)
            {
                dataWords = lines[line]
                    .Split(Delim)
                    .Select(P => P.Trim(separator))
                    //.Where(p => !string.IsNullOrWhiteSpace(p))
                    .ToArray();

                DataRow row = myTable.NewRow();

                int index = 0;
                while (index < dataWords.Length) // loop through each line add the data to the row
                {
                    foreach (string data in dataWords)
                    {
                        row[index] = data;
                        index++;
                    }
                }
                myTable.Rows.Add(row);
            }
        }
    }
}
