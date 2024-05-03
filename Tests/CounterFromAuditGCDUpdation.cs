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
using System.Threading.Tasks;

namespace SwaggerWebAPI
{
    class CounterFromAuditGCDUpdation
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();

        static string ABF_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetCounterPartyIdFromAuditGcdIdUpdation']/div[1]/h3/span[1]/a",
                         ABF_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetCounterPartyIdFromAuditGcdIdUpdation_content']/form/div[2]/input",
                         ABF_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetCounterPartyIdFromAuditGcdIdUpdation_content']/div[2]/div[4]/pre",
                        ABF_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetCounterPartyIdFromAuditGcdIdUpdation_content']/div[2]/div[3]/pre/code",
                        Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetCounterPartyIdFromAuditGcdIdUpdation_content']/form/table/tbody/tr/td/input[@name='Username']",
                        ABF_input = "//div[@id='DatabaseTopic_DatabaseTopic_GetCounterPartyIdFromAuditGcdIdUpdation_content']/form/table/tbody/tr/td/input[@class='parameter required']";

        public CounterFromAuditGCDUpdation(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_ABF => Driver.FindElement(By.XPath(ABF_url));
        IWebElement button_ABF => Driver.FindElement(By.XPath(ABF_Try));
        IWebElement input_ABF => Driver.FindElement(By.XPath(ABF_input));
        IWebElement body_ABF => Driver.FindElement(By.XPath(ABF_body));
        IWebElement res_ABF => Driver.FindElement(By.XPath(ABF_code));
        IWebElement un => Driver.FindElement(By.XPath(Username));

        public void GetCounterFromAudit()
        {

            test = ExtReport.CreateTest("TC_018_CounterFromAuditpartyGCDUpdation").Info("CounterAudit Test has Started");

            string val1 = InputExcelAPI.GetCellData("WebAPI", 12, 3);

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);*/
            //v//ar row = sheet.GetRow(11);
            //var val = row.GetCell(0).StringCellValue;
            //var val = row.GetCell(0).StringCellValue;
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
            //Thread.Sleep(2000);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(ABF_url)));

            url_ABF.Click();

            test.Log(Status.Info, "CounerAudit selected");
            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(ABF_input)));

            input_ABF.SendKeys(val1);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            string input = string.Format(username);
            un.SendKeys(input);

            test.Log(Status.Info, "counterpartyId " + val1 + " has entered");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(ABF_Try)));

            button_ABF.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(ABF_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_ABF);
            action.Perform();

            string resBody = body_ABF.Text;
            string resCode = res_ABF.Text;

            //test.Log(Status.Info, resBody);
            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "CounterAudit Response is " + resCode);

                if (resBody=="true")
                {
                    test.Log(Status.Pass, "CounterAudit Response body contains " + "True");
                    test.Log(Status.Pass, "Test Passed!");
                }
                else
                {
                    test.Log(Status.Fail, "AuditGCD Response is failed!! not contain " + "False");
                    Assert.Fail("Fixingsource Response is false!");

                }
            }
            else
            {
                test.Log(Status.Fail, "AuditGCD Response is " + resCode);
                Assert.Fail("Fixingsource Response is ", resCode);
            }

            url_ABF.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

        }
    }
}


