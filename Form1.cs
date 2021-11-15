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

        private void button1_Click(object sender, EventArgs e)
        {
            LoadCSVs();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveAsCsv();
        }

        private void LoadExcel()
        {
            var csvpath = @"F:\src\mypumaconfigurator\PUMAConfigurator\PUMA_Bill_of_Materials.xls";
        }

        private void LoadCSVs()
        {
            var csvpath = Form1.csvpath;
            var files = Directory.EnumerateFiles(csvpath).ToArray();
            try
            {
                foreach (var file in files)
                {
                    var PUMA = new PumaObject(file);
                    using var parser = new TextFieldParser(file);
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters("\t");
                    int i = 0;

                    if (PUMA.CsvType == CsvType.stl)
                    {
                        PumaObject innerPO = null;

                        while (!parser.EndOfData)
                        {
                            string[] fields = parser.ReadFields();
                            if (i != 0
                                && fields.Length > 1
                                && fields[1] == "Time_Hr")
                            {
                                innerPO = new PumaObject() { ID = fields[0], CsvType = CsvType.stl };
                                PUMA.BOMSTRUCT.Add(innerPO);
                            }
                            if (i != 0
                                && innerPO != null
                                && fields[1].EndsWith("stl"))
                            {
                                innerPO.BOMSTRUCT.Add(new PumaObject() { ID = fields[0], CsvType = CsvType.stl });
                            }

                            i++;
                        }
                    }
                    if (PUMA.CsvType == CsvType.config)
                    {
                        // read 00 for name
                        // then everything in 2nd col that is not empty and after 'Module'
                        // thats the struct
                        bool modules = false;
                        while (!parser.EndOfData)
                        {
                            //Processing row
                            string[] fields = parser.ReadFields();
                            if (fields.Length <= 2)
                            {
                                break;
                            }
                            if (i == 0)
                            {
                                PUMA.Descr = fields[0];
                            }
                            if (i == 1)
                            {
                                PUMA.Descr2 = fields[0];
                            }
                            if (modules)
                            {
                                var X = fields[0];
                                PUMA.BOMSTRUCT.Add(new PumaObject() { ID = X });
                            }
                            if (fields[1] == "Module")
                            {
                                modules = true;
                            }

                            i++;
                        }
                    }
                    if (PUMA.CsvType == CsvType.nonprintedpart)
                    {
                    }
                    if (PUMA.CsvType == CsvType.md)
                    {
                        PumaObject innerPO = null;
                        bool model = false;
                        bool npp = false;
                        while (!parser.EndOfData)
                        {
                            string[] fields = parser.ReadFields();
                            if (i == 0)
                            {
                                PUMA.Descr = fields[0];
                            }
                            if (npp)
                            {
                                int num = 0;
                                if (int.TryParse(fields[5], out num))
                                {
                                    innerPO = new PumaObject() { ID = fields[0], CsvType = CsvType.nonprintedpart, Quantity = num };
                                    PUMA.BOMSTRUCT.Add(innerPO);
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
                                int num = 0;
                                if (fields.Length < 5)
                                    continue;
                                if (int.TryParse(fields[5], out num))
                                {
                                    innerPO = new PumaObject() { ID = fields[0], CsvType = CsvType.stl, Quantity = num };
                                    PUMA.BOMSTRUCT.Add(innerPO);
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

                                //innerPO = new PumaObject() { ID = fields[0], CsvType = CsvType.stl };
                                //PUMA.BOMSTRUCT.Add(innerPO);
                            }
                            //if (i != 0
                            //    && innerPO != null
                            //    && fields[1].EndsWith("stl"))
                            //{
                            //    innerPO.BOMSTRUCT.Add(new PumaObject() { ID = fields[0], CsvType = CsvType.stl });
                            //}

                            i++;
                        }
                    }
                }
            }
            catch (System.IndexOutOfRangeException e)
            {
            }
        }

        public static void SaveAsCsv()
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
    }
}