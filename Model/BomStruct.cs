using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PUMAConfigurator.Model
{
    public class BomStruct
    {
        public static List<BomStruct> BomStructs { get; set; } = new List<BomStruct>();

        public static void AddBomStruct(BomStruct bomstruct)
        {
            BomStructs.Add(bomstruct);
        }

        public static List<BomStruct> GetItemTree(string ID)
        {
            var structure = BomStructs.Where(m=>m.ParentID == ID).ToList();
            var substruct = structure.SelectMany(m => GetItemTree(m.ChildID)).ToList();
            if (substruct.Count == 0)
                return structure;
            else
                return substruct;
        }

        public string ParentID { get; set; }
        public string ChildID { get; set; }
        public int Qt { get; set; }

        public override string ToString()
        {
            return $"{ParentID}\t{Qt}\t{ChildID}";
        }
    }
}
