using System;
using System.Windows.Forms;
using System.Net.Http;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;
using System.Diagnostics;
using System.Security.Policy;

using Excel = Microsoft.Office.Interop.Excel;

using Duck.OfficeAutomationModule.Office;
using Duck.OfficeAutomationModule.Selenium;

namespace Duck
{
    public partial class Form1 : Form
    {
        ExcelFileHandler mExcelHandler = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void form1_Shown(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Executable Files|*.exe|All Files|*.*"; 
                openFileDialog.Title = "Select ChromeDriver Executable"; 
                openFileDialog.DefaultExt = "exe"; 
                openFileDialog.AddExtension = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    try
                    {
                        ProteinInfoCrawler.SetChromeDriverPath(filePath);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex);
                    }
                }
                else
                {
                    MessageBox.Show(
                        "Chrome Driver를 설정해야 사용할 수 있습니다.", "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                }
            }

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
            percentageNumericUpDown.Value = 0.1m;
        }

        private void openFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (mExcelHandler != null)
            {
                mExcelHandler.Close();
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx|All Files|*.*";
                openFileDialog.Title = "Save an Excel File";
                openFileDialog.DefaultExt = "xlsx";
                openFileDialog.AddExtension = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
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
