using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SectionBox
{
    public partial class LoadExcel : Form
    {
        private Autodesk.Revit.UI.UIApplication m_app;

        public LoadExcel(Autodesk.Revit.UI.UIApplication app)
        {
            InitializeComponent();

            m_app = app;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var mainCommand = new Command();

            OpenFileDialog theDialog = new OpenFileDialog();
                theDialog.Title = "Open Excel File";
                theDialog.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";
                theDialog.InitialDirectory = @"Libraries\Documents";
                //if the user clicked OK and not cancel
                if (theDialog.ShowDialog() == DialogResult.OK)
                {
                    string filename = theDialog.FileName;
                    mainCommand.loadExcelSheet(filename);
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Please, select the Excel file containing the section views names");
                }
        }
    }
}
