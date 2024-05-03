
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
    class AddedByFromAuditGCDUpdation {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;

        string username = User.getEncodedUserName();

        static string ABF_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetAddedByFromAuditGcdIdUpdation']/div[1]/h3/span[1]/a",
                         ABF_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetAddedByFromAuditGcdIdUpdation_content']/form/div[2]/input",
                        ABF_input = "//div[@id='DatabaseTopic_DatabaseTopic_GetAddedByFromAuditGcdIdUpdation_content']/form/table/tbody/tr/td/input[@class='parameter required']",
                         ABF_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetAddedByFromAuditGcdIdUpdation_content']/div[2]/div[4]/pre",
                        ABF_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetAddedByFromAuditGcdIdUpdation_content']/div[2]/div[3]/pre/code",
                        Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetAddedByFromAuditGcdIdUpdation_content']/form/table/tbody/tr/td/input[@name='Username']";

        public AddedByFromAuditGCDUpdation(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement body_ABF => Driver.FindElement(By.XPath(ABF_body));
        IWebElement res_ABF => Driver.FindElement(By.XPath(ABF_code));

        public void GetAddeByFrom()
        {

            test = ExtReport.CreateTest("TC_017_AddedByFromAuditGCDUpdation").Info("Test has Started");

            string val1 = InputExcelAPI.GetCellData("WebAPI", 12, 3);

            string val = ValidationExcelAPI.GetCellData("DatabaseTopic", 20, 3);

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(20);
             var val = row.GetCell(3).StringCellValue;*/
            //var val = row.GetCell(0).StringCellValue;
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
            // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(ABF_url)));

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(ABF_url)));

            url_ABF.Click();

            test.Log(Status.Info, "AuditGCD selected");
            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(8);

//            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            string input = string.Format(username);
            un.SendKeys(input);

            input_ABF.SendKeys(val1);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(ABF_Try)));

            button_ABF.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
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
                test.Log(Status.Pass, "AuditGCD Response is " + resCode);

                if (resBody.Contains(val))
                {
                    test.Log(Status.Pass, "AuditGCD Response body contains " + val);
                    test.Log(Status.Pass, "Test Passed!");
                }
                else
                {
                    test.Log(Status.Fail, "AuditGCD Response is failed!! not contain " + val);
                    Assert.Fail("Fixingsource Response is failed! "+resBody);
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


