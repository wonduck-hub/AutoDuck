using HtmlAgilityPack;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using static System.Net.WebRequestMethods;
using System.Diagnostics;

namespace Duck
{
    public partial class Form1 : Form
    {
        private string mUrl = "https://www.seoultech.ac.kr/index.jsp";
        public Form1()
        {
            InitializeComponent();
        }

        private async Task readWebPage(string url)
        {
            // HttpClient를 사용하여 웹 페이지 내용 가져오기
            HttpClient client = new HttpClient();
            HttpResponseMessage response = await client.GetAsync(url);
            string pageContents = await response.Content.ReadAsStringAsync();

            // HtmlAgilityPack을 사용하여 HTML 문서 파싱
            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(pageContents);

            // CSS 선택자를 사용하여 데이터 추출
            HtmlNodeCollection nodes = document.DocumentNode.SelectNodes("//span[@class='title']"); // CSS 선택자 사용
            foreach (var node in nodes)
            {
                Debug.WriteLine(node.InnerText);
            }
        }


        private async void generateBtn_Click(object sender, EventArgs e)
        {
            await readWebPage(mUrl);
        }
    }
}
