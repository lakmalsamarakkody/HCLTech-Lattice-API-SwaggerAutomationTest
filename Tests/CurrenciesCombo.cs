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
    class CurrenciesCombo
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_AddCurrenciesCombo']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_AddCurrenciesCombo_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_AddCurrenciesCombo_content']/div[2]/div[4]/pre",
                         inputD = "//div[@id='DatabaseTopic_DatabaseTopic_AddCurrenciesCombo_content']/form/table/tbody/tr/td/textarea[@name='lstresult']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_AddCurrenciesCombo_content']/div[2]/div[3]/pre/code";

        public CurrenciesCombo(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement data => Driver.FindElement(By.XPath(inputD));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void Get_CurrenciesCombo()
        {
            test = ExtReport.CreateTest("TS_056_GetCurrenciesCombo").Info("Test Started");

            string value = InputExcelAPI.GetCellData("WebAPI", 38, 3);

            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 48, 3);
            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(5);
            var row = sheet.GetRow(58);
            var value = row.GetCell(3).StringCellValue;

            Thread.Sleep(2000);*/

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "CurrenciesCombo selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(inputD)));

            string input = string.Format(@"['{1}','{0}']", username, value);
            data.SendKeys(input);

            test.Log(Status.Info, "Data : "+value+ " entered");

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
                test.Log(Status.Pass, "CurrenciesCombo Response is " + resCode);

                if (resBody.Contains(value))
                {
                    test.Log(Status.Pass, "CurrenciesCombo Response contains " + value);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else if (resBody == "no content")
                {
                    test.Log(Status.Warning, "CurrenciesCombo Response doesn't have any values");
                    Assert.Warn("CurrenciesCombo Response has " + resBody);
                }

                else
                {
                    test.Log(Status.Fail, "CurrenciesCombo Response is Failed!");
                    Assert.Fail("CurrenciesCombo Response is " + resBody);

                }
            }
            else
            {
                test.Log(Status.Fail, "CurrenciesCombo Response is " + resCode);
                Assert.Fail("CurrenciesCombo Response is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
