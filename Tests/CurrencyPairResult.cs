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
    class CurrencyPairResult
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetCurrencyPairResults']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetCurrencyPairResults_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetCurrencyPairResults_content']/div[2]/div[4]/pre",
                         userName = "//*[@id='DatabaseTopic_DatabaseTopic_GetCurrencyPairResults_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetCurrencyPairResults_content']/div[2]/div[3]/pre/code";


        public CurrencyPairResult(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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


        public void GetCurrencyPairResult()
        {

            test = ExtReport.CreateTest("TS_050_GetCurrencyPairResult").Info("Test Started");

            //string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 68, 3);
            //string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 68, 4);

            SqlCommand cmd = new SqlCommand("select * from CcyPair where ID = (select max(ID) from CcyPair)", con);

            //cmd.Parameters.AddWithValue("@Code", value);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string code;
            string Id;

            code = Convert.ToString(reader[1]).Trim();
            Id = Convert.ToString(reader[0]);

            reader.Close();
            con.Close();

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "GetCurrencyPairResult selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));

            string input = string.Format(username);
            un.SendKeys(input);

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
                test.Log(Status.Pass, "GetCurrencyPairResult Response is " + resCode);

                if (resBody.Contains(Id)&& resBody.Contains(code))
                {
                    test.Log(Status.Pass, "GetCurrencyPairResult Response body contains latest ccypair ID " +Id+ " & Code "+code);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetCurrencyPairResult Response is Failed");
                    Assert.Fail("GetCurrencyPairResult Response is Failed!!");
                }
            }
            else
            {
                test.Log(Status.Fail, "GetCurrencyPairResult Response is " + resCode);
                Assert.Fail("GetCurrencyPairResult Response is " + resCode);
            }

            url_lattise.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
