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
    class GetSetting
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='SystemSettings_SystemSettings_GetSettings']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='SystemSettings_SystemSettings_GetSettings_content']/form/div[2]/input",
                         lattise_code = "//*[@id='SystemSettings_SystemSettings_GetSettings_content']/div[2]/div[4]/pre",
                         userName = "//div[@id='SystemSettings_SystemSettings_GetSettings_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='SystemSettings_SystemSettings_GetSettings_content']/div[2]/div[3]/pre/code";


        public GetSetting(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement un => Driver.FindElement(By.XPath(userName));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetSystemSetting()
        {
            test = ExtReport.CreateTest("TS_083_GetSettings").Info("Test Started");

            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 135, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 135, 4);
            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(5);
            var row = sheet.GetRow(131);
            var value = row.GetCell(3).StringCellValue;
            var val2 = row.GetCell(4).StringCellValue;

            Thread.Sleep(2000);*/
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
            url_lattise.Click();
            test.Log(Status.Info, "GetSystemSetting selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            string input = string.Format(username);
            un.SendKeys(username);

            //IWebElement result = Driver.FindElement(By.XPath("//textarea[@Name='userDetails']"));

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
                test.Log(Status.Pass, "GetSystemSetting Response is " + resCode);

                if (resBody.Contains(val1) && resBody.Contains(val2))
                {
                    test.Log(Status.Pass, "GetSystemSetting Response contains Name: " + val1 + " & status: "+val2);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetSystemSetting Response is Failed!");
                    Assert.Fail("GetSystemSetting Response is Failed!! ");
                }
            }
            else
            {
                test.Log(Status.Fail, "GetSystemSetting Response is " + resCode);
                Assert.Fail("GetSystemSetting Response code is " + resCode);
            }

            url_lattise.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
