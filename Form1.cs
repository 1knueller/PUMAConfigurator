using PUMAConfigurator.Domain;
using PUMAConfigurator.Model;
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
        public Form1()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
        }

        private void excel2csvtoolStripButton1_Click(object sender, EventArgs e)
        {
            FileDataHandler.Excel2csv();
        }

        private void loadcsvtoolStripButton2_Click(object sender, EventArgs e)
        {
            FileDataHandler.LoadCSVs();
            this.dataGridView1.DataSource = PumaObject.Pumas;
        }

        private void rowstatechanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            if (e.StateChanged == DataGridViewElementStates.Selected)
            {
                var R = e.Row;
                if (R.DataBoundItem is PumaObject po)
                {
                    var itemtree = BomStruct.GetItemTree(po.ID);

                    var itemtree2 = itemtree.OrderBy(strct => PumaObject.Pumas.FirstOrDefault(m => m.ID == strct.ChildID).CsvType)
                        .ThenBy(m=>m.ChildID);

                    var STR = string.Join(Environment.NewLine, itemtree2.Select(m => m.ToString()));
                    this.textBox1.Text = STR;
                }
            }
        }
    }
}