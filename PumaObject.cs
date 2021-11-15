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
        public static List<PumaObject> Pumas { get; set; } = new List<PumaObject>();

        public PumaObject()
        {
            Pumas.Add(this);
        }

        public PumaObject(string file)
        {
            ID = Path.GetFileName(file);
            var type = ID.Substring(0, 1).ToUpperInvariant();

            if (!CSVTYPEDIC.DIC.ContainsKey(type))
                return;

            CsvType = CSVTYPEDIC.DIC[type];
            Pumas.Add(this);
        }

        public string csvpath { get; set; }

        public string ID { get; set; }

        public CsvType CsvType { get; set; }

        public string Descr { get; set; }
        public string Descr2 { get; set; }

        public double Weight { get; set; }
    }
}
