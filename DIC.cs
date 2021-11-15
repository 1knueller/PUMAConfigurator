
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PUMAConfigurator
{
    public static class CSVTYPEDIC
    {
        public static Dictionary<string, CsvType> DIC = new Dictionary<string, CsvType>()
        {
            { "L",CsvType.stl},
            { "C",CsvType.config},
            { "M",CsvType.md},
            { "N",CsvType.nonprintedpart},
        };
    }
}
