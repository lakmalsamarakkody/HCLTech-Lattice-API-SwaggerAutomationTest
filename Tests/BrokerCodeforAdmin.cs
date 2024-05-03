using AventStack.ExtentReports;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SwaggerWebAPI.Libs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace SwaggerWebAPI
{
    class BrokerCodeforAdmin
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();

        static string url = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerBrokerCodeResultsForAdmin']/div[1]/h3/span[1]/a",
                         Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerBrokerCodeResultsForAdmin_content']/form/div[2]/input",
                         code = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerBrokerCodeResultsForAdmin_content']/div[2]/div[4]/pre",
                        Bname=  "//div[@id='DatabaseTopic_DatabaseTopic_GetBrokerBrokerCodeResultsForAdmin_content']/form/table/tbody/tr/td/input[@class='parameter required']",
                         Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetBrokerBrokerCodeResultsForAdmin_content']/form/table/tbody/tr/td/input[@name='Username']",
                        body = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerBrokerCodeResultsForAdmin_content']/div[2]/div[3]/pre/code";
        public BrokerCodeforAdmin(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_admin => Driver.FindElement(By.XPath(url));
        IWebElement button_admin => Driver.FindElement(By.XPath(Try));
        IWebElement body_admin => Driver.FindElement(By.XPath(body));
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement name => Driver.FindElement(By.XPath(Bname));
        IWebElement res_admin => Driver.FindElement(By.XPath(code));


        public void GetBrokerCodeforAdmin()
        {

            test = ExtReport.CreateTest("Get_BrokerCodeforAdmin").Info("BrokerCodeforAdmin Test has Started");

            string val = InputExcelAPI.GetCellData("WebAPI", 19, 3);
            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 36, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 36, 4);

            // string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            /*XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(5);
            var row = sheet.GetRow(36);
            var val = row.GetCell(3).NumericCellValue.ToString();
            var val2 = row.GetCell(4).NumericCellValue.ToString();
            //var val = row.GetCell(0).StringCellValue;
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
            Thread.Sleep(2000);*/
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(url)));

            url_admin.Click();

          //  Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

            test.Log(Status.Info, "BrokerCodeforAdmin selected");

            name.SendKeys(val);
            test.Log(Status.Info, "Broker Name : nom3 entered");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            string input = string.Format(username);
            un.SendKeys(input);

            button_admin.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_admin);
            action.Perform();

            string resBody = body_admin.Text;
            string resCode = res_admin.Text;

            //test.Log(Status.Info, resBody);
            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "BrokerCodeforAdmin Response is " + resCode);

                if (resBody.Contains(val))
                {
                    test.Log(Status.Pass, "BrokerCodeforAdmin Response body contain BrokerID " +val1+ " & postingID "+val2);
                    test.Log(Status.Pass, "Test Passed!");
                }
                else if (resBody=="[]")
                {
                    test.Log(Status.Warning, "BrokerCodeforAdmin Response body doesn't contain any values");
                    Assert.Warn("BrokerCodeforAdmin Response is "+ resBody);
                }
                else
                {
                    test.Log(Status.Fail, "BrokerCodeforAdmin Response is failed!! not contain " + val + ", Instead it contains " + resBody);
                    Assert.Fail("BrokerCodeforAdmin Response is ", resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "BrokerCodeforAdmin Response is " + resCode);
                Assert.Fail("BrokerCodeforAdmin Response is "+ resCode);
            }

            url_admin.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);

        }
    }
}

