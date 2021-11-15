using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PUMAConfigurator
{
    public class PumaObject
    {
        public PumaObject()
        {

        }
        public PumaObject(string file)
        {
            ID = Path.GetFileName(file);
            var type = ID.Substring(0, 1).ToUpperInvariant();

            CsvType = CSVTYPEDIC.DIC[type];
        }

        public string csvpath { get; set; }

        public string ID { get; set; }

        public CsvType CsvType { get; set; }

        public int Quantity { get; set; }

        public string Descr { get; set; }
        public string Descr2 { get; set; }

        public List<PumaObject> BOMSTRUCT = new List<PumaObject>();
    }
}
