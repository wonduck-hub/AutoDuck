using HtmlAgilityPack;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using static System.Net.WebRequestMethods;
using System.Diagnostics;
using System.Security.Policy;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Duck
{
    public partial class Form1 : Form
    {
        private string mUrl = "https://www.seoultech.ac.kr/index.jsp";
        private string mHtmlClassName = "title";

        public Form1()
        {
            InitializeComponent();
        }

        private async Task<HtmlNodeCollection> readWebPage(string url)
        {
            Debug.Assert(url != null);
            Debug.Assert(url != String.Empty);

            // HttpClient를 사용하여 웹 페이지 내용 가져오기
            HttpClient client = new HttpClient();
            HttpResponseMessage response = await client.GetAsync(url);

            string pageContents = await response.Content.ReadAsStringAsync();

            // HtmlAgilityPack을 사용하여 HTML 문서 파싱
            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(pageContents);

            // CSS 선택자를 사용하여 데이터 추출
            HtmlNodeCollection nodes = document.DocumentNode.SelectNodes($"//span[@class='{mHtmlClassName}']");

            return nodes;
        }

        private void makeXlsxFile(string filePath, HtmlNodeCollection nodes)
        {
            Debug.Assert(filePath != null);
            Debug.Assert(filePath != String.Empty);
            Debug.Assert(nodes != null);

            // 새로운 Excel 파일 생성
            using (SpreadsheetDocument spreadsheetDocument = 
                    SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                // WorkbookPart 추가
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // WorksheetPart 추가
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // 시트 데이터 추가
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // 첫 번째 행 추가
                for (int i = 0; i < nodes.Count; ++i)
                {
                    HtmlNode node = nodes[i];
                    Row row1 = new Row() { RowIndex = (uint)(i + 1) };
                    Cell cellB1 = new Cell() { CellReference = $"B{ i + 1 }",
                        CellValue = new CellValue(node.InnerText.ToString()), DataType = CellValues.String };
                    row1.Append(cellB1);
                    sheetData.Append(row1);
                }

                // 시트 및 문서 저장
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.AppendChild(new Sheets()).AppendChild(new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "DataSheet"
                });
                workbookPart.Workbook.Save();

            }
        }
            private async void generateBtn_Click(object sender, EventArgs e)
        {
            HtmlNodeCollection nodes = await readWebPage(mUrl);

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files|*.xlsx|All Files|*.*";
                saveFileDialog.Title = "Save an Excel File";
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.AddExtension = true;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    // save xlsx file
                    makeXlsxFile(filePath, nodes);

                    MessageBox.Show($"File will be saved to: {filePath}", "Save",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

        }
    }
}
