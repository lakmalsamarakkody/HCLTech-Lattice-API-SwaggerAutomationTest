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
    class BrokerCodeIdResult
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerCodeIdResult']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerCodeIdResult_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerCodeIdResult_content']/div[2]/div[4]/pre",
                         uname = "//div[@id='DatabaseTopic_DatabaseTopic_GetBrokerCodeIdResult_content']/form/table/tbody/tr/td/input[@name='Username']",
                         code = "//div[@id='DatabaseTopic_DatabaseTopic_GetBrokerCodeIdResult_content']/form/table/tbody/tr/td/input[@name='brokerCode']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerCodeIdResult_content']/div[2]/div[3]/pre/code";


        public BrokerCodeIdResult(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement inputCode => Driver.FindElement(By.XPath(code));
        IWebElement un => Driver.FindElement(By.XPath(uname));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetBrokerCodeIdResult()
        {
            test = ExtReport.CreateTest("TC_027_BrokerCodeIdResult").Info("Test Started");

            string val = InputExcelAPI.GetCellData("WebAPI", 22, 3);
            //string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 32, 3);

            SqlCommand cmd = new SqlCommand("select * from BrokerCode where Code = @Code", con);
            cmd.Parameters.AddWithValue("@Code", val);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string id;
            string bcode;
            bcode = Convert.ToString(reader[1]);
            id = Convert.ToString(reader[0]);

            reader.Close();
            con.Close();

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(26);
             var value = row.GetCell(3).NumericCellValue.ToString();

             Thread.Sleep(2000);*/

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));


            url_lattise.Click();

            test.Log(Status.Info, "BrokerCodeIdResult selected");
            // Thread.Sleep(4000);


            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(code)));

            inputCode.SendKeys(val);
            test.Log(Status.Info, "Broker code : nom3 entered");

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(uname)));

            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(lattise_Try)));

            button_lattise.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            // Thread.Sleep(5000);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(lattise_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_lattise);
            action.Perform();

            string resBody = body_lattise.Text;
            string resCode = res_Code.Text;


            test.Log(Status.Info, "Verifying Value....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "BrokerCodeIdResult Response is " + resCode);

                if (resBody.Contains(id))
                {
                    test.Log(Status.Pass, "BrokerCodeIdResult Response contains ID: " + id);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else if (resBody.Equals("-1"))
                {
                    test.Log(Status.Warning, "BrokerCodeIdResult Response dosen't make any sense");
                    Assert.Warn("BrokerCodeIdResult Response dosen't make any sense");

                }
                else 
                {
                    test.Log(Status.Fail, "BrokerCodeIdResult Response is Faile! Instead in contains " + resBody);
                    Assert.Fail("BrokerCodeIdResult Response is "+ resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "BrokerCodeIdResult Response is " + resCode);
                Assert.Fail("BrokerCodeIdResult Response is "+ resCode);
            }

            url_lattise.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
