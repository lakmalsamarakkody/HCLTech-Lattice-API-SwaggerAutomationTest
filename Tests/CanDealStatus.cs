using System;
using OpenQA.Selenium;
using System.Threading;
using AventStack.ExtentReports;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SwaggerWebAPI.Libs;
using NUnit.Framework;
using System.Data.SqlClient;

namespace SwaggerWebAPI
{
    class CanDealStatus
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string status_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetCanDealStatus']/div[1]/h3/span[1]/a",
                         status_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetCanDealStatus_content']/form/div[2]/input",
                         Res_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetCanDealStatus_content']/div[2]/div[4]/pre",
                         lstresult = "//*[@id='DatabaseTopic_DatabaseTopic_GetCanDealStatus_content']/form/table/tbody/tr/td/textarea[@name='lstresult']",
                         status_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetCanDealStatus_content']/div[2]/div[3]/pre/code";

        public CanDealStatus(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement statusField => Driver.FindElement(By.XPath(lstresult));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));


        public void GetDealstatus()
        {

            test = ExtReport.CreateTest("TS_009_GetCanDealStatus").Info("Deal Status Test has Started");

            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 3, 3);

            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(0);
            var row = sheet.GetRow(4);
            var val1 = row.GetCell(0).NumericCellValue.ToString();*/

            //WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(5));

            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);
            // wait.Until(SeleniumExtras.WaitHelpers.
            //   WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(status_url)));

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(status_url)));

            url_status.Click();

            test.Log(Status.Info, "Deal Status selected");
            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

            string input = string.Format(@"['{1}','{0}']", username,val1);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lstresult)));
            //date.SendKeys(@"['2023-04-05T19:58:02.6575227+05:30',"+username+"]");
            statusField.SendKeys(input);
            //string input2 = string.Format(@"['CORP\l_samarakkody'],'{0}']", username);
            //  string input3 = string.Format(@"['CORP\\l_samarakkody'],'{0}']", username);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(status_Try)));

            button_status.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");
            // string input2 = string.Format(@"['CORP\\l_samarakkody'],'{0}']", username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(status_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_status);
            action.Perform();


            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);
            // Thread.Sleep(4000);

            string resBody = body_status.Text;
            string resCode = res_Code.Text;


            test.Log(Status.Info, "Verifying Values....");

            test.Log(Status.Pass, "Deal Status Response is " + resBody);


            if (resCode == "200")
            {
                test.Log(Status.Pass, "Deal Status Response code is " + resCode);
               

                 if (resBody=="true")
                { 
                     test.Log(Status.Pass, "Deal Status Response  " + resBody);
                     test.Log(Status.Pass, "Deal Status Test is Pass");
                 }
                 else
                 {
                     test.Log(Status.Fail, "Deal Status Response is failed!!");
                    Assert.Fail("Response code " +resBody);

                }
            }
            else
            {
                test.Log(Status.Fail, "Deal Status failed Response code is " + resCode);
                Assert.Fail("Response code " + resCode);

            }

            url_status.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
