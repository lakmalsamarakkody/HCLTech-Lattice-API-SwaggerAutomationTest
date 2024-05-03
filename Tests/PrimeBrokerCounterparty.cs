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
using System.Threading.Tasks;

namespace SwaggerWebAPI
{
    class PrimeBrokerCounterparty
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string PBC_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetAllPrimeBrokerCounterparties']/div[1]/h3/span[1]/a",
                         PBC_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetAllPrimeBrokerCounterparties_content']/form/div[2]/input",
                         PBC_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetAllPrimeBrokerCounterparties_content']/div[2]/div[4]/pre",
                        Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetAllPrimeBrokerCounterparties_content']/form/table/tbody/tr/td/input[@name='Username']",
                        PBC_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetAllPrimeBrokerCounterparties_content']/div[2]/div[3]/pre/code";

        public PrimeBrokerCounterparty(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }        

        IWebElement url_PBC => Driver.FindElement(By.XPath(PBC_url));
        IWebElement button_PBC => Driver.FindElement(By.XPath(PBC_Try));
        IWebElement body_PBC => Driver.FindElement(By.XPath(PBC_body));
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement res_PBC => Driver.FindElement(By.XPath(PBC_code));

        public void GetPrimeBrokerCounter()
        {
            test = ExtReport.CreateTest("TC_014_PrimeBrokerCounter").Info("PrimeBrokerCounter Test has Started");
            string val = ValidationExcelAPI.GetCellData("DatabaseTopic", 12, 3);

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";
             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(12);
             //var val = row.GetCell(0).NumericCellValue.ToString();
             var val = row.GetCell(3).StringCellValue;
             //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);
             Thread.Sleep(2000);*/
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(PBC_url)));

            url_PBC.Click();

            test.Log(Status.Info, "PrimeBrokerCounter selected");
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(1000);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(PBC_Try)));

            string input = string.Format(username);
            un.SendKeys(input);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            button_PBC.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(PBC_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_PBC);
            action.Perform();

            string resBody = body_PBC.Text;
            string resCode = res_PBC.Text;

            test.Log(Status.Info, "Verifying Response values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "PrimeBrokerCounter Response is " + resCode);

                if (resBody.Contains(val))
                {
                    test.Log(Status.Pass, "PrimeBrokerCounter Response body contains " + val);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "PrimeBrokerCounter Response is failed!! not contain " + val );
                    Assert.Fail("PrimeBrokerCounter Response contain " + resBody);

                }
            }
            else
            {
                test.Log(Status.Fail, "Failed!! BrokerCodes Response is " + resCode);
                Assert.Fail("PrimeBrokerCounter Response is " + resCode);

            }

            WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(5));

            url_PBC.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
