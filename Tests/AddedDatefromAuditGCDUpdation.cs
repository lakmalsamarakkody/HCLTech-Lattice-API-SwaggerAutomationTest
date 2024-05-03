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
    class AddedDatefromAuditGCDUpdation
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;

        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();


        static string Agcd_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetAddedDateFromAuditGcdIdUpdation']/div[1]/h3/span[1]/a",
                         Agcd_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetAddedDateFromAuditGcdIdUpdation_content']/form/div[2]/input",
                         Agcd_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetAddedDateFromAuditGcdIdUpdation_content']/div[2]/div[4]/pre",
                        Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetAddedDateFromAuditGcdIdUpdation_content']/form/table/tbody/tr/td/input[@name='Username']",
                        Agcd_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetAddedDateFromAuditGcdIdUpdation_content']/div[2]/div[3]/pre/code";

        public AddedDatefromAuditGCDUpdation(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

            IWebElement url_Agcd => Driver.FindElement(By.XPath(Agcd_url));
        IWebElement button_Agcd => Driver.FindElement(By.XPath(Agcd_Try));
        IWebElement body_Agcd => Driver.FindElement(By.XPath(Agcd_body));
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement statusField => Driver.FindElement(By.Name("counterpartyId"));
        IWebElement res_Agcd => Driver.FindElement(By.XPath(Agcd_code));


        public void GetAuditGCD()
        {

            test = ExtReport.CreateTest("TC_016_AddedDatefromAuditGCDUpdation").Info("AuditGCD Test has Started");

            string val = ValidationExcelAPI.GetCellData("DatabaseTopic", 11, 3);
            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 24, 3);

            SqlCommand cmd = new SqlCommand("select * from AuditGcdIdUpdation where CounterPartyId= 1374 ", con);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            string maxDate;
            string counterId;

            reader.Read();

            maxDate = reader[4].ToString();
            counterId = reader[1].ToString();

            reader.Close();
            con.Close();
            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(1);
             var val = row.GetCell(21).StringCellValue;
             //var val = row.GetCell(0).StringCellValue;
             //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
             Thread.Sleep(2000);*/
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Agcd_url)));

            url_Agcd.Click();

           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);

            test.Log(Status.Info, "AddedDatefromAuditGCDUpdation selected");
            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            statusField.SendKeys(counterId);
            //un.Click();
            string input = string.Format(username);
            un.SendKeys(input);
            test.Log(Status.Info, "counterpartyId " + counterId+" has entered");

            button_Agcd.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Agcd_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_Agcd);
            action.Perform();

            //string date = maxDate.ToString("yyyy-mm-dd");

            string resBody = body_Agcd.Text;
            string resCode = res_Agcd.Text;

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
                    Assert.Fail("AuditGCD Response contain " + val); 
                }
            }
            else
            {
                test.Log(Status.Fail, "AuditGCD Response is " + resCode);
                Assert.Fail("AuditGCD Response contain " + resCode);

            }

            url_Agcd.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

        }
    }
}
