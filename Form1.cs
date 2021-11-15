using ExcelDataReader;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PUMAConfigurator
{
    internal partial class Form1 : Form
    {
        private const string csvpath = "csvs";
        private const string excelfilename = "PUMA_Bill_of_Materials.xls";

        public List<PumaObject> PUMAS { get; private set; }

        public Form1()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
        }

        string x = "SNMC";

        private void LoadCSVs()
        {
            var csvpath = Form1.csvpath;
            var files = Directory.EnumerateFiles(csvpath)
                .OrderBy(fp => x.IndexOf(Path.GetFileName(fp)
                    .Substring(0, 1).ToUpperInvariant())).ToArray();
            PumaObject.Pumas.Clear();

            foreach (var file in files)
            {
                var filePuma = new PumaObject(Path.GetFileNameWithoutExtension(file));
                using var parser = new TextFieldParser(file);
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters("\t");
                int i = 0;

                if (filePuma.CsvType == CsvType.STL)
                {
                    PumaObject stlPartCollection = null;

                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        if (i != 0
                            && fields.Length > 1
                            && fields[1] == "Time_Hr")
                        {
                            stlPartCollection = new PumaObject() { ID = fields[0], CsvType = CsvType.STL };
                        }
                        if (i != 0
                            && stlPartCollection != null
                            && fields[0].EndsWith("stl"))
                        {
                            new PumaObject() { ID = fields[0], CsvType = CsvType.STL };

                            BomStruct.BomStructs.Add(
                                new BomStruct()
                                {
                                    ParentID = stlPartCollection.ID,
                                    ChildID = fields[0]
                                });
                        }

                        i++;
                    }
                }
                if (filePuma.CsvType == CsvType.Configuration)
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
                                ChildID = "MD_"+fields[1],
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
                if (filePuma.CsvType == CsvType.NonPrintedPart)
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
                                CsvType = CsvType.NonPrintedPart,
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
                if (filePuma.CsvType == CsvType.Module)
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

            this.dataGridView1.DataSource = PumaObject.Pumas;
        }

        public static void Excel2csv()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using var stream = new FileStream(excelfilename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IExcelDataReader reader = null;
            if (excelfilename.EndsWith(".xls"))
            {
                reader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (excelfilename.EndsWith(".xlsx"))
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
            Directory.CreateDirectory(csvpath);

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

                File.WriteAllText(Path.Combine(csvpath, $"{table.TableName}.csv"), csvContent.ToString());
            }
        }

        private void excel2csvtoolStripButton1_Click(object sender, EventArgs e)
        {
            Excel2csv();
        }

        private void loadcsvtoolStripButton2_Click(object sender, EventArgs e)
        {
            LoadCSVs();
        }

        private void rowstatechanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            if (e.StateChanged == DataGridViewElementStates.Selected)
            {
                var R = e.Row;
                if (R.DataBoundItem is PumaObject po)
                {
                    var Y = BomStruct.GetItemTree(po.ID);
                    var STR = string.Join(Environment.NewLine, Y.Select(m => m.ToString()));
                    this.textBox1.Text = STR;
                }
            }
        }
    }
}