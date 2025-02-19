using System;
using System.Windows.Forms;
using System.Net.Http;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;
using System.Diagnostics;
using System.Security.Policy;

using Excel = Microsoft.Office.Interop.Excel;

using Duck.OfficeAutomationModule.Office;
using Duck.OfficeAutomationModule.RestApi;
using System.Xml;

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
            if (!ExcelFileHandler.IsExcelInstalled())
            {
                MessageBox.Show(
                        "MS Excel이 설치되어 있지 않습니다.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }

            disableAllControl();
        }

        #region Uniprot API
        private async void proteinInfoButton_Click(object sender, EventArgs e)
        {
            string uniProtSerialNum = mExcelHandler.GetActiveCellValue();
            string uniProtXML = await UniProtApi.GetProteinDataOrNullAsync(uniProtSerialNum);

            if (uniProtXML != null)
            {
                proteinInfoTextBox.Text = String.Empty;

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(uniProtXML);

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("u", "http://uniprot.org/uniprot");

                XmlNode nameNode = doc.SelectSingleNode("//*[local-name()='entry']");
                if (nameNode != null)
                {
                    string entryName = 
                        nameNode.SelectSingleNode("*[local-name()='name']").InnerText;

                    proteinInfoTextBox.AppendText($"Entry Name: {entryName}{Environment.NewLine}");
                    proteinInfoTextBox.AppendText(Environment.NewLine);
                }
                else
                {
                    proteinInfoTextBox.AppendText($"Entry Name 없음{Environment.NewLine}");
                }

                XmlNode geneNode = doc.SelectSingleNode("//*[local-name()='gene']");
                if (nameNode != null)
                {
                    string geneName =
                        geneNode.SelectSingleNode("*[local-name()='name' and @type='primary']").InnerText;

                    proteinInfoTextBox.AppendText($"Gene Name: {geneName}{Environment.NewLine}");
                    proteinInfoTextBox.AppendText(Environment.NewLine);
                }
                else
                {
                    proteinInfoTextBox.AppendText($"Gene Name 없음{Environment.NewLine}");
                }

                XmlNode proteinRecommendedNameNode = 
                    doc.SelectSingleNode("//*[local-name()='protein']/*[local-name()='recommendedName']");
                if (nameNode != null)
                {
                    string proteinFullName =
                        proteinRecommendedNameNode.SelectSingleNode("*[local-name()='fullName']").InnerText;

                    proteinInfoTextBox.AppendText(
                        $"Protein Full Name: {proteinFullName}{Environment.NewLine}");
                    proteinInfoTextBox.AppendText(Environment.NewLine);
                }
                else
                {
                    proteinInfoTextBox.AppendText($"Gene Name 없음{Environment.NewLine}");
                }

                XmlNode organismNode = doc.SelectSingleNode("//*[local-name()='organism']");
                if (organismNode != null)
                {
                    string scientificName = 
                        organismNode.SelectSingleNode("*[local-name()='name' and @type='scientific']").InnerText;
                    string commonName = 
                        organismNode.SelectSingleNode("*[local-name()='name' and @type='common']").InnerText;
                    string taxonomyId = 
                        organismNode.SelectSingleNode(
                            "*[local-name()='dbReference' and @type='NCBI Taxonomy']").Attributes["id"].Value;

                    proteinInfoTextBox.AppendText($"Scientific Name: {scientificName}{Environment.NewLine}");
                    proteinInfoTextBox.AppendText($"Common Name: {commonName}{Environment.NewLine}");
                    proteinInfoTextBox.AppendText($"Taxonomy ID: {taxonomyId}{Environment.NewLine}");
                    proteinInfoTextBox.AppendText(Environment.NewLine);

                    proteinInfoTextBox.AppendText($"Lineage:{Environment.NewLine}");
                    XmlNodeList lineageNodes = 
                        organismNode.SelectNodes("*[local-name()='lineage']/*[local-name()='taxon']");
                    foreach (XmlNode taxon in lineageNodes)
                    {
                        proteinInfoTextBox.AppendText($"{taxon.InnerText}{Environment.NewLine}");
                    }
                    proteinInfoTextBox.AppendText(Environment.NewLine);
                }
                else
                {
                    proteinInfoTextBox.AppendText($"Organism 없음{Environment.NewLine}");
                }
            }
            else
            {
                proteinInfoTextBox.Text = $"error!";
            }
        }
        #endregion

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
            proteinInfoButton.Enabled = false;
            proteinInfoTextBox.Enabled = false;
        }

        private void enableAllControl()
        {
            showExcelWindowCheckBox.Enabled = true;
            worksheetsComboBox.Enabled = true;
            runButton.Enabled = true;
            saveFileToolStripMenuItem.Enabled = true;
            percentageNumericUpDown.Enabled = true;
            proteinInfoButton.Enabled = true;
            proteinInfoTextBox.Enabled = true;
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
            Stopwatch watch = new Stopwatch();

            disableAllControl();
            this.Cursor = Cursors.WaitCursor;

            watch.Start();
            bool isSucess = mExcelHandler.CetsaMsRun(worksheetsComboBox.SelectedIndex + 1, percentageNumericUpDown.Value);
            watch.Stop();

            this.Cursor = Cursors.Default;
            enableAllControl();
            initControl();

            if (isSucess)
            {
                MessageBox.Show("sucess!\n경과 시간: " + watch.ElapsedMilliseconds + "ms", "Sucess",
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
