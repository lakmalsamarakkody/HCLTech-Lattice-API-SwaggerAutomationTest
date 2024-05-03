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
    class GetBrokerCodeResults
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerCodeResults']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerCodeResults_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerCodeResults_content']/div[2]/div[4]/pre",
                         uname = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerCodeResults_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerCodeResults_content']/div[2]/div[3]/pre/code";

        public GetBrokerCodeResults(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement un => Driver.FindElement(By.XPath(uname));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));

        public void Get_BrokerCodeResult()
        {
            test = ExtReport.CreateTest("GetBrokerCodeResult").Info("Test Started");

            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 34, 3);
            //string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 34, 4);

            SqlCommand cmd = new SqlCommand("select * from BrokerCode where ID = @ID", con);
            cmd.Parameters.AddWithValue("@ID", val1);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string id;
            string code;
            id = Convert.ToString(reader[0]);
            code = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();

            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(5);
            var row = sheet.GetRow(26);
            var value = row.GetCell(3).NumericCellValue.ToString();

            Thread.Sleep(2000);*/

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "GetBrokerCodeResult selected");
            // Thread.Sleep(4000);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(uname)));
            string input = string.Format(username);
            un.SendKeys(input);
            //test.Log(Status.Info, "Code : 360JA entered");

            button_lattise.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            //Thread.Sleep(5000);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(lattise_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_lattise);
            action.Perform();

            string resBody = body_lattise.Text;
            string resCode = res_Code.Text;

            test.Log(Status.Info, "Verifying Value....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GetBrokerCodeResult Response is " + resCode);

                if (resBody.Contains(id))
                {
                    test.Log(Status.Pass, "GetBrokerCodeResult Response body contains BrokerCodeId " + id + " & BrokerCode " + code);
                    test.Log(Status.Pass, "Test is Pass");
                }

                else
                {
                    test.Log(Status.Fail, "GetBrokerCodeResult Response is Failed");
                    Assert.Fail("GetBrokerCodeResult Response is failed!");
                }
            }
            else
            {
                test.Log(Status.Fail, "GetBrokerCodeResult Response is " + resCode);
                Assert.Fail("GetBrokerCodeResult Response is " + resCode);
            }

            url_lattise.Click();
          //  Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
