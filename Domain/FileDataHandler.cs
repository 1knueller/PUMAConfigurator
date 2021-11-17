using ExcelDataReader;
using Microsoft.VisualBasic.FileIO;
using PUMAConfigurator.Model;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PUMAConfigurator.Domain
{
    public static class FileDataHandler
    {
        private const string _excelfilename = "PUMA_Bill_of_Materials.xls";
        private const string _csvpath = "csvs";
        private const string _orderOfReadingCSVTypes = "CMNS";

        public static void LoadCSVs()
        {
            var files = Directory.EnumerateFiles(_csvpath)
                .OrderByDescending(fp => _orderOfReadingCSVTypes.IndexOf(Path.GetFileName(fp)
                    .Substring(0, 1).ToUpperInvariant())).ToArray();

            PumaObject.Pumas.Clear();
            BomStruct.BomStructs.Clear();

            foreach (var file in files)
            {
                var filePuma = new PumaObject(Path.GetFileNameWithoutExtension(file));
                using var parser = new TextFieldParser(file);
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters("\t");
                int i = 0;

                if (filePuma.CsvType == ECsvType.STL)
                {
                    PumaObject stlPartCollection = null;

                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();

                        if (i != 0
                            && fields.Length > 1
                            && fields[1] == "Time_Hr")
                        {
                            stlPartCollection = new PumaObject() { ID = fields[0], CsvType = ECsvType.STL };
                        }
                        else if (i != 0
                            && stlPartCollection != null
                            //&& fields[0].EndsWith("stl") // sadly not all stl files in the list end with .stl
                            )
                        {
                            var stlfilename = fields[0];

                            new PumaObject() { ID = stlfilename, CsvType = ECsvType.STL };

                            BomStruct.BomStructs.Add(
                                new BomStruct()
                                {
                                    ParentID = stlPartCollection.ID,
                                    ChildID = stlfilename
                                });
                        }

                        i++;
                    }
                }
                if (filePuma.CsvType == ECsvType.Configuration)
                {
                    bool modules = false;
                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        if (fields.Length <= 2)
                        {
                            break;
                        }
                        if (i == 0)
                        {
                            filePuma.Descr = fields[0];
                        }
                        if (i == 1)
                        {
                            filePuma.Descr2 = fields[0];
                        }
                        if (modules && !string.IsNullOrWhiteSpace(fields[1]))
                        {
                            var partId = fields[1];

                            BomStruct.AddBomStruct(new BomStruct()
                            {
                                ParentID = filePuma.ID,
                                ChildID = "MD_" + fields[1],
                                Qt = 1,
                            });
                        }
                        if (!modules && fields[1] == "Module")
                        {
                            modules = true;
                        }

                        i++;
                    }
                }
                if (filePuma.CsvType == ECsvType.NonPrintedPart)
                {
                    PumaObject innerPO = null;
                    bool contentStarted = false;
                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        if (contentStarted)
                        {
                            if (fields.Length < 3)
                                continue;
                            double.TryParse(fields[2], out double weight);

                            innerPO = new PumaObject()
                            {
                                ID = fields[0],
                                Descr = fields[1],
                                CsvType = ECsvType.NonPrintedPart,
                                Weight = weight,
                            };
                        }
                        if (!contentStarted
                            && fields[0] == "Length of threaded shaft")
                        {
                            contentStarted = true;
                        }

                        i++;
                    }
                }
                if (filePuma.CsvType == ECsvType.Module)
                {
                    bool model = false;
                    bool npp = false;
                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        if (i == 0)
                        {
                            filePuma.Descr = fields[0];
                        }
                        if (npp)
                        {
                            int num = 0;
                            if (int.TryParse(fields[5], out num))
                            {
                                BomStruct.AddBomStruct(new BomStruct()
                                {
                                    ParentID = filePuma.ID,
                                    ChildID = fields[0],
                                    Qt = num,
                                });
                            }
                            else
                                continue;
                        }
                        if (!npp
                            && fields.Length > 0
                            && fields[0] == "Non-Printed Parts (NPP)")
                        {
                            npp = true;
                        }
                        if (!npp
                            && model)
                        {
                            if (fields.Length < 5)
                                continue;
                            if (int.TryParse(fields[5], out int num))
                            {
                                BomStruct.AddBomStruct(new BomStruct()
                                {
                                    ParentID = filePuma.ID,
                                    ChildID = fields[0],
                                    Qt = num,
                                });
                            }
                            else
                                continue;
                        }

                        if (!npp
                            && model == false
                            && i != 0
                            && fields.Length > 0
                            && fields[0] == "Model")
                        {
                            model = true;
                        }

                        i++;
                    }
                }
            }

            PumaObject.SortPumas();
        }

        public static void Excel2csv()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using var stream = new FileStream(_excelfilename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IExcelDataReader reader = null;
            if (_excelfilename.EndsWith(".xls"))
            {
                reader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (_excelfilename.EndsWith(".xlsx"))
            {
                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            if (reader == null)
                return;

            var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = false
                }
            });
            Directory.CreateDirectory(_csvpath);

            var sb = new StringBuilder();
            for (int itable = 0; itable < ds.Tables.Count; itable++)
            {
                var csvContent = sb.Clear();
                var table = ds.Tables[itable];
                for (int irows = 0; irows < table.Rows.Count; irows++)
                {
                    var arr = new List<string>();
                    for (int icol = 0; icol < table.Columns.Count; icol++)
                    {
                        arr.Add(table.Rows[irows][icol].ToString());
                    }
                    csvContent.AppendJoin("\t", arr);
                    csvContent.AppendLine();
                }

                File.WriteAllText(Path.Combine(_csvpath, $"{table.TableName}.csv"), csvContent.ToString());
            }
        }
    }
}