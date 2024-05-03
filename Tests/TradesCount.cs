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
    class TradesCount
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();


        static string lattise_url = "//*[@id='Trade_Trade_GetTradesCount']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='Trade_Trade_GetTradesCount_content']/form/div[2]/input",
                         lattise_code = "//*[@id='Trade_Trade_GetTradesCount_content']/div[2]/div[4]/pre",
                         deal = "//div[@id='Trade_Trade_GetTradesCount_content']/form/table/tbody/tr/td/input[@name='dealId']",
                         userName = "//div[@id='Trade_Trade_GetTradesCount_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='Trade_Trade_GetTradesCount_content']/div[2]/div[3]/pre/code";


        public TradesCount(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement id => Driver.FindElement(By.XPath(deal));
        IWebElement un => Driver.FindElement(By.XPath(userName));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetTradesCount()
        {

            test = ExtReport.CreateTest("TS_088_GetTradesCount").Info("Test Started");

            string value = InputExcelAPI.GetCellData("WebAPI", 54, 3);
            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 151, 3);

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(119);
             var value = row.GetCell(3).NumericCellValue.ToString();
           Thread.Sleep(2000);*/
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
            url_lattise.Click();
            test.Log(Status.Info, "GetTradesCount selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(deal)));
            id.SendKeys(value);
            test.Log(Status.Info, "DealID : "+value+" has entered");

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


            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GetTradesCount Response is " + resCode);

                if (resBody.Equals(value))
                {
                    test.Log(Status.Pass, "GetTradesCount Response contains Trades count : " + resBody);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else if (resBody == "null")
                {
                    test.Log(Status.Warning, "GetTradesCount Response is Null!");
                }
                
                else if (resBody == "0")
                {
                    test.Log(Status.Warning, "GetTradesCount Response is 0, All Trades has been removed for this Deal");
                    Assert.Warn("GetTradesCount Response is " + resBody);
                }
                else
                {
                    test.Log(Status.Fail, "GetTradesCount Response has Failed!!");
                    Assert.Fail("GetTradesCount Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "GetTradesCount Response is " + resCode);
                Assert.Fail("GetTradesCount Response code is " + resCode);
            }

            url_lattise.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
