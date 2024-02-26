using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using Dev_Test.Util;


namespace Dev_Test.Model
{
    class Electoral: Utils
    {
        public override void ExtractDataExcludingHeader(char Delim, string[] lines)
        {
            char separator = '"';

            ExtractDataUsingSeparator(Delim, lines, separator);
        }   
    }
}
