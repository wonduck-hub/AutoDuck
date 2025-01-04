using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools.V129.SystemInfo;

namespace Duck.OfficeAutomationModule.Selenium
{
    static public class ProteinInfoCrawler
    {
        static public string chromeDriverPath;

        static public void SetChromeDriverPath(string path)
        {
            chromeDriverPath = path;
        }
        static public void GetInfo(string name)
        {
            Debug.Assert(chromeDriverPath != null);
            Debug.Assert(name != string.Empty);

            ChromeOptions options = new ChromeOptions(); 
            //options.AddArgument("--headless"); // 헤드리스 모드 설정
            ChromeDriver driver = null;
            try
            {
                driver = new ChromeDriver(chromeDriverPath, options);
                driver.Navigate().GoToUrl("https://www.uniprot.org/");

                // TODO: 웹에서 자동화할 내용 작성
            }
            catch (WebDriverException ex)
            {
                Console.WriteLine("WebDriverException caught: " + ex.Message);
                Console.WriteLine("Chrome 브라우저 또는 ChromeDriver를 확인하세요.");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                if (driver != null)
                {
                    driver.Quit();
                }
            }
        }
    }
}
