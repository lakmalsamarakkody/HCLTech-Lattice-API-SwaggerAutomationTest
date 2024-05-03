using AventStack.ExtentReports;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using SwaggerWebAPI.Libs;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace SwaggerWebAPI
{
    class LocalRecieved
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string status_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetRecordsWithStatusAsLocalReceived']/div[1]/h3/span[2]/a",
                         status_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetRecordsWithStatusAsLocalReceived_content']/form/div[2]/input",
                         Res_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetRecordsWithStatusAsLocalReceived_content']/div[2]/div[4]/pre",
                         Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetRecordsWithStatusAsLocalReceived_content']/form/table/tbody/tr/td/input[@name='Username']",
                        status_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetRecordsWithStatusAsLocalReceived_content']/div[2]/div[3]/pre/code";

        public LocalRecieved(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_status => Driver.FindElement(By.XPath(status_url));
        IWebElement button_status => Driver.FindElement(By.XPath(status_Try));
        IWebElement body_status => Driver.FindElement(By.XPath(status_body));
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));

        public void GetLocalRecieved()
        {

            test = ExtReport.CreateTest("TS_012_GetCanDealStatus").Info("Deal Status Test has Started");

            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(0);
            var row = sheet.GetRow(4);
            var val1 = row.GetCell(0).NumericCellValue.ToString();*/

            //WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(5));

            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
            // wait.Until(SeleniumExtras.WaitHelpers.
           // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(status_url)));

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(status_url)));

            url_status.Click();

            test.Log(Status.Info, "GetRecordsWithStatusAsLocalReceived selected");
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            string input = string.Format(username);
            un.SendKeys(input);

            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(status_Try)));

            button_status.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(status_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_status);
            action.Perform();

            //Thread.Sleep(4000);

            string resBody = body_status.Text;
            string resCode = res_Code.Text;

            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "LocalReceived Response code is " + resCode);

                if (resBody == "true")
                {
                    test.Log(Status.Pass, "LocalReceived Response  " + resBody);
                    test.Log(Status.Pass, "Deal Status Test is Pass");
                }

                else if(resBody.Contains("null"))
                {
                    test.Log(Status.Warning, "LocalReceived Response is Null");
                }
                else
                {
                    test.Log(Status.Fail, "LocalReceived Response is failed!!");
                }
            }
            else
            {
                test.Log(Status.Fail, "LocalReceived failed Response code is " + resCode);
                Assert.Fail("Fixingsource Response is "+resCode);

            }

            url_status.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
