using System;
using System.Net.Http;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;
using System.Diagnostics;
using System.Security.Policy;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

using OfficeFileHandler;

namespace Duck
{
    public partial class Form1 : Form
    {
        OfficeFileHandler.ExcelFileHandler mExcelHandler = null;

        public Form1()
        {
            InitializeComponent();

            disableAllControl();
        }

        #region Load Close
        private void form1_Load(object sender, EventArgs e)
        {
        }

        private void form1_Close(object sender, FormClosedEventArgs e)
        {
            if (mExcelHandler != null)
            {
                mExcelHandler.Dispose();
            }
        }
        #endregion

        #region disable enable
        private void disableAllControl()
        {
            showExcelWindowCheckBox.Enabled = false;
            worksheetsComboBox.Enabled = false;
            runButton.Enabled = false;
            saveFileToolStripMenuItem.Enabled = false;
            percentageNumericUpDown.Enabled = false;
        }

        private void enableAllControl()
        {
            showExcelWindowCheckBox.Enabled = true;
            worksheetsComboBox.Enabled = true;
            runButton.Enabled = true;
            saveFileToolStripMenuItem.Enabled = true;
            percentageNumericUpDown.Enabled = true;
        }
        #endregion

        private void initControl()
        {
            Excel.Sheets ws = mExcelHandler.GetSheets();

            worksheetsComboBox.Items.Clear();
            foreach (Excel.Worksheet sheet in ws)
            {
                worksheetsComboBox.Items.Add(sheet.Name);
            }
            worksheetsComboBox.SelectedIndex = 0;
            percentageNumericUpDown.Value = 0.0m;
        }

        private void openFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (mExcelHandler != null)
            {
                mExcelHandler.Close();
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files|*.xlsx|All Files|*.*";
                saveFileDialog.Title = "Save an Excel File";
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.AddExtension = true;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;
                    try
                    {
                        mExcelHandler = new ExcelFileHandler(filePath);

                        enableAllControl();
                        initControl();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex);
                    }
                }
            }
        }

        private void showExcelWindowCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Debug.Assert(mExcelHandler != null);

            System.Windows.Forms.CheckBox checkBox = sender as System.Windows.Forms.CheckBox;

            if (checkBox.Checked)
            {
                mExcelHandler.SetVisible(true);
            }
            else
            {
                mExcelHandler.SetVisible(false);
            }
        }

        private void runButton_Click(object sender, EventArgs e)
        {
            disableAllControl();
            this.Cursor = Cursors.WaitCursor;

            bool isSucess = mExcelHandler.MsCetsaRun(worksheetsComboBox.SelectedIndex + 1, percentageNumericUpDown.Value);

            this.Cursor = Cursors.Default;
            enableAllControl();
            initControl();

            if (isSucess)
            {
                MessageBox.Show("sucess!", "Sucess",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("can't found valid table", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mExcelHandler.Save();
        }
    }
}
