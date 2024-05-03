using AventStack.ExtentReports;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
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
    class ExpireTrades
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='Trade_Trade_GetExpireTrades']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='Trade_Trade_GetExpireTrades_content']/form/div[2]/input",
                         lattise_code = "//*[@id='Trade_Trade_GetExpireTrades_content']/div[2]/div[4]/pre",
                         userName = "//div[@id='Trade_Trade_GetExpireTrades_content']/form/table/tbody/tr/td/input[@name='Username']",
                         expire = "//div[@id='Trade_Trade_GetExpireTrades_content']/form/table/tbody/tr/td/input[@name='expiryFilter']",
                        lattise_body = "//*[@id='Trade_Trade_GetExpireTrades_content']/div[2]/div[3]/pre/code";


        public ExpireTrades(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_lattise => Driver.FindElement(By.XPath(lattise_url));
        IWebElement button_lattise => Driver.FindElement(By.XPath(lattise_Try));
        IWebElement body_lattise => Driver.FindElement(By.XPath(lattise_body));
        IWebElement result => Driver.FindElement(By.XPath(expire));
        IWebElement un => Driver.FindElement(By.XPath(userName));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetExpireTrades()
        {
            test = ExtReport.CreateTest("TS_086_GetExpireTrades").Info("Test Started");

            string value = InputExcelAPI.GetCellData("WebAPI", 50, 3);

            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 143, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 143, 4);
            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(5);
            var row = sheet.GetRow(111);
            var value = row.GetCell(3).NumericCellValue.ToString();
            var val2 = row.GetCell(4).NumericCellValue.ToString();

            Thread.Sleep(2000);*/

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
            url_lattise.Click();
            test.Log(Status.Info, "GetExpireTrades selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(expire)));
            result.SendKeys(value);
            test.Log(Status.Info, "Expire Date: "+value+ " entered");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            string input = string.Format(username);
            un.SendKeys(username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_Try)));
            button_lattise.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lattise_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_lattise);
            action.Perform();

            string resBody = body_lattise.Text;
            string resCode = res_Code.Text;


            test.Log(Status.Info, "Verifying Value....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GetExpireTrades Response is " + resCode);

                if (resBody.Contains(val1) && resBody.Contains(val2))
                {
                    test.Log(Status.Pass, "GetExpireTrades Response contains DealID: " + val1 + " , TradeID: " + val2);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetExpireTrades Response is Failed!");
                    Assert.Fail("GetExpireTrades Response is Failed!!");
                }
            }
            else
            {
                test.Log(Status.Fail, "GetExpireTrades Response is " + resCode);
                Assert.Fail("GetExpireTrades Response code is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
