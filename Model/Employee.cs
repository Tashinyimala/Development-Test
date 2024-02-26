using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using System.Data;

namespace Dev_Test
{
    class Employee
    {
        private int _ID;
        private string _FirstName;
        private string _lastName;
        private string _Address1;
        private string _Address2;
        private string _Address3;
        private string _Address4;
        private string _Address5;

        private string _FilePathname;
        private BindingList<Employee> _employeeList = new BindingList<Employee>();
        
        public string[] headerLables = new List<String>().ToArray();

        public BindingList<Employee> employeeList {
            get => _employeeList;

            set => _employeeList = value;
        }

        public int ID
        {
            get => _ID;
            set  => _ID = value;
        }

        public string FirstName
        {
            get => _FirstName;
            set => _FirstName = value;
        }

        public string LastName
        {
            get => _lastName;
            set => _lastName = value;
        }

        public string Address1
        {
            get => _Address1;
            set => _Address1 = value;
        }

        public string Address2
        {
            get => _Address2;
            set => _Address2 = value;
        }

        public string Address3
        {
            get => _Address3;
            set => _Address3 = value;
        }

        public string Address4
        {
            get => _Address4;
            set => _Address4 = value;
        }

        public string Address5
        {
            get => _Address5;
            set => _Address5 = value;
        }

        public string FilePathname
        {
            get => _FilePathname;
            set => _FilePathname = value;
        }

        public void ReadCommaSepratedFle(char Delim)
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

        private void ExtractDataExcludingHeader(char Delim, string[] lines)
        {
            // Extracting Data in employeeList
            for (int line = 1; line < lines.Length; line++)
            {
                string[] dataWords = lines[line].Split(Delim);

                employeeList.Add(new Employee
                {
                    ID = int.Parse(dataWords[0]),
                    FirstName = dataWords[1],
                    LastName = dataWords[2],
                    Address1 = dataWords[3],
                    Address2 = dataWords[4],
                    Address3 = dataWords[5],
                    Address4 = dataWords[6],
                    Address5 = dataWords[7],
                });
            }
        }

        private void ExtractHeader(char Delim, string[] lines)
        {
            // Extracting Header from the file
            string firstLine = lines[0];
            //headerLables = firstLine.Split(Delim);
            headerLables = firstLine
                            .Split(Delim)
                            .Select(p => p.Trim())
                            .Where(p => !string.IsNullOrWhiteSpace(p))
                            .ToArray();
        }
    }
}
