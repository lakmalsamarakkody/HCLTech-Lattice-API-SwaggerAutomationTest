using AventStack.ExtentReports;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
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
    class TraderBrokerCodeResult
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();


        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetTraderBrokerCodeResults']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetTraderBrokerCodeResults_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetTraderBrokerCodeResults_content']/div[2]/div[4]/pre",
                         input = "//div[@id='DatabaseTopic_DatabaseTopic_GetTraderBrokerCodeResults_content']/form/table/tbody/tr/td/textarea[@name='resultlst']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetTraderBrokerCodeResults_content']/div[2]/div[3]/pre/code";


        public TraderBrokerCodeResult(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement result => Driver.FindElement(By.XPath(input));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetTraderBrokerCodeIdResult()
        {
            test = ExtReport.CreateTest("Get_TraderBrokerCodeIdResult").Info("Test Started");

            string inputD = InputExcelAPI.GetCellData("WebAPI", 24, 3);
            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 36, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 36, 4);

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(38);
             var value = row.GetCell(3).NumericCellValue.ToString();
             var val2 = row.GetCell(4).StringCellValue;

             Thread.Sleep(2000);*/

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "TraderBrokerCodeIdResult selected");
            // Thread.Sleep(4000);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(input)));

            string data = string.Format("['{1}','{0}']",username,inputD);
            result.SendKeys(data);
            test.Log(Status.Info, "Broker details entered");

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(lattise_Try)));

            button_lattise.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(lattise_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_lattise);
            action.Perform();

            string resBody = body_lattise.Text;
            string resCode = res_Code.Text;


            test.Log(Status.Info, "Verifying Value....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "TraderBrokerCodeIdResult Response is " + resCode);

                if (resBody.Contains(val1) && resBody.Contains(val2))
                {
                    test.Log(Status.Pass, "TraderBrokerCodeIdResult Response contains BrokerCodeID " + val1 + " TraderCode " + val2);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "TraderBrokerCodeIdResult Response is Faile! Instead in contains " + resBody);
                    Assert.Fail("TraderBrokerCodeIdResult Response is Failed!" );
                }
            }
            else
            {
                test.Log(Status.Fail, "TraderBrokerCodeIdResult Response is " + resCode);
                Assert.Fail("TraderBrokerCodeIdResult Response is " + resCode);
            }

            url_lattise.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
