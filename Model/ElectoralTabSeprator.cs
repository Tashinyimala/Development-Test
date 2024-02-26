using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dev_Test.Util;

namespace Dev_Test.Model
{
    class ElectoralTabSeprator : Utils
    {
        public override void ExtractDataExcludingHeader(char Delim, string[] lines)
        {
            char separator = '\t';
            ExtractDataUsingSeparator(Delim, lines, separator);

        }
    }
}
