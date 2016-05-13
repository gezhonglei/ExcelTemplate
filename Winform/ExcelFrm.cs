using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExportTemplate.Export;

namespace Winform
{
    public partial class ExcelFrm : Form
    {
        Dictionary<string, string> comboValue = null;
        public ExcelFrm()
        {
            InitializeComponent();

            Init();
        }

        private void Init()
        {
            this.StartPosition = FormStartPosition.CenterScreen;

            //ExportConfig config = ExportConfig.NewInstance(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelTemplate\\ExportConfig.xml"));
            //cmbProductType.Items.AddRange(config.ProductTypes.Select(p => p.Name).ToArray());
            comboValue = ExportMain.GetProductDescription();
            //cmbProductType.Items.AddRange(ExportMain.GetProductList().ToArray());
            cmbProductType.Items.AddRange(comboValue.Values.ToArray());
            txtFileName.Text = Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "test.xlsx");
        }

        private void btnSaveAsExcel_Click(object sender, EventArgs e)
        {
            if (cmbProductType.SelectedItem == null)
            {
                return;
            }

            this.Enabled = false;
            int index = cmbProductType.SelectedIndex;
            //LY.BM.Studio.Comm.Export.ExportTest.Test2(cmbProductType.SelectedItem.ToString(), txtFileName.Text);
            ExportTest.Test2(comboValue.Keys.ToArray()[index], txtFileName.Text);
            this.Enabled = true;
        }

        private void btnSaveName_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel 97-2003 Workbook(*.xls)|*.xls|Excel Workbook(*.xlsx)|*.xlsx";
            saveFileDialog.FilterIndex = rdbtnXls.Checked ? 1 : 2;
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtFileName.Text = saveFileDialog.FileName;
                if (new FileInfo(txtFileName.Text).Extension == ".xls")
                {
                    rdbtnXls.Checked = true;
                }
                else
                {
                    rdbtnXlsx.Checked = true;
                }
            }
        }

        private void ExcelFormat_CheckedChanged(object sender, EventArgs e)
        {
            string name = txtFileName.Text;
            txtFileName.Text = name.Replace(new FileInfo(name).Extension, rdbtnXls.Checked ? ".xls" : ".xlsx");
        }
    }
}
