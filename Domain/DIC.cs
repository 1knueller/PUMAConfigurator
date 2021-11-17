using PUMAConfigurator.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PUMAConfigurator.Domain
{
    public static class CSVTYPEDIC
    {
        public static Dictionary<string, ECsvType> DIC = new Dictionary<string, ECsvType>()
        {
            { "L",ECsvType.STL},
            { "C",ECsvType.Configuration},
            { "M",ECsvType.Module},
            { "N",ECsvType.NonPrintedPart},
            { "F",ECsvType.FilamentDensity},
        };
    }
}
