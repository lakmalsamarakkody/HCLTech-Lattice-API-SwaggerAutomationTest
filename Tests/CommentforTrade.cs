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
    class CommenetforTrade
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetCommentForTrades']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetCommentForTrades_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetCommentForTrades_content']/div[2]/div[4]/pre",
                         trade = "//div[@id='DatabaseTopic_DatabaseTopic_GetCommentForTrades_content']/form/table/tbody/tr/td/textarea[@name='TradeIdsAndCounterPartyIds']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetCommentForTrades_content']/div[2]/div[3]/pre/code";

        public CommenetforTrade(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement tradeCounter => Driver.FindElement(By.XPath(trade));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetCommenetforTrade()
        {

            test = ExtReport.CreateTest("TS_047_CommenetforTrade").Info("Test Started");

            string value = InputExcelAPI.GetCellData("WebAPI", 35, 3);

            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 28, 3);
            

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(69);
             var val2 = row.GetCell(4).NumericCellValue.ToString();
             var val3 = row.GetCell(5).NumericCellValue.ToString();
             var value = row.GetCell(3).StringCellValue;
             Thread.Sleep(2000);*/

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

            test.Log(Status.Info, "CommenetforTrade selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(trade)));

          //  string input = string.Format(@"{{1},'{0}}", username, value);
            string input = "{" + value + "'" + username + "'}";
            tradeCounter.SendKeys(input);

            test.Log(Status.Info, "Date Time has entered");

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
                test.Log(Status.Pass, "CommenetforTrade Response is " + resCode);

                if (resBody.Contains(val1))
                {
                    test.Log(Status.Pass, "CommenetforTrade Response body contains " + val1);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else if(resBody.Equals("{}"))
                {
                    test.Log(Status.Warning, "CommenetforTrade Response doesn't contain any data!");
                    Assert.Warn("CommenetforTrade Response doesn't contain any data!");
                }else
                {
                    test.Log(Status.Fail, "CommenetforTrade Response is Failed");
                    Assert.Fail("CommenetforTrade Response is Failed!");
                }
            }
            else
            {
                test.Log(Status.Fail, "CommenetforTrade Response is " + resCode);
                Assert.Fail("CommenetforTrade Response is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(1000);

        }
    }
}
