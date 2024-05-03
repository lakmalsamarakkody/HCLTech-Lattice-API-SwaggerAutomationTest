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
    class RangeOfTrades
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='Trade_Trade_GetRangeOfTrades']/div[1]/h3/span[1]/a",
                         lattise_Try = "//*[@id='Trade_Trade_GetRangeOfTrades_content']/form/div[2]/input",
                         lattise_code = "//*[@id='Trade_Trade_GetRangeOfTrades_content']/div[2]/div[4]/pre",
                         inputD = "//div[@id='Trade_Trade_GetRangeOfTrades_content']/form/table/tbody/tr/td/textarea[@name='filterparamlst']",
                        lattise_body = "//*[@id='Trade_Trade_GetRangeOfTrades_content']/div[2]/div[3]/pre/code";

        public RangeOfTrades(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement details => Driver.FindElement(By.XPath(inputD));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetRangeOfTrades()
        {
            test = ExtReport.CreateTest("TS_085_GetRangeOfTrades").Info("Test Started");

            string value = InputExcelAPI.GetCellData("WebAPI", 49, 3);

            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 159, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 159, 4);

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(123);
             var value = row.GetCell(3).NumericCellValue.ToString();
             var val2 = row.GetCell(4).StringCellValue;

             Thread.Sleep(2000);*/

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
            url_lattise.Click();
            test.Log(Status.Info, "GetRangeOfTrades selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(inputD)));
            // details.SendKeys(@"['CORP\\l_samarakkody','DESKTOP-NJK2OUL','2.7.6.7']");
            string input = "[" + value + "'" + username + "']";
            //string input = string.Format($"[{value},'{username}']");

            details.SendKeys(input);
            test.Log(Status.Info, "User Details entered");

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
                test.Log(Status.Pass, "GetRangeOfTrades Response is " + resCode);

                if (resBody.Contains(val1) && resBody.Contains(val2))
                {
                    test.Log(Status.Pass, "GetRangeOfTrades Response contains user ID: " + val1 + " & Name: " + val2);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetRangeOfTrades Response is Failed!");
                    Assert.Fail("GetRangeOfTrades Response is Failed!!");
                }
            }
            else
            {
                test.Log(Status.Fail, "GetRangeOfTrades Response is " + resCode);
                Assert.Fail("GetRangeOfTrades Response code is " + resCode);
            }

            url_lattise.Click();
            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);
        }
    }
}
