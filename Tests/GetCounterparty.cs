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
    class GetCounterparty
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string url = "//*[@id='DatabaseTopic_DatabaseTopic_GetCounterparty']/div[1]/h3/span[2]/a",
                         Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetCounterparty_content']/form/div[2]/input",
                         code = "//*[@id='DatabaseTopic_DatabaseTopic_GetCounterparty_content']/div[2]/div[4]/pre",
                          name = "//div[@id='DatabaseTopic_DatabaseTopic_GetCounterparty_content']/form/table/tbody/tr/td/input[@class='parameter required']",
                         uname = "//*[@id='DatabaseTopic_DatabaseTopic_GetCounterparty_content']/form/table/tbody/tr/td/input[@name='Username']",
                         body = "//*[@id='DatabaseTopic_DatabaseTopic_GetCounterparty_content']/div[2]/div[3]/pre/code";
        public GetCounterparty(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_dup => Driver.FindElement(By.XPath(url));
        IWebElement button_dup => Driver.FindElement(By.XPath(Try));
        IWebElement body_dup => Driver.FindElement(By.XPath(body));
        IWebElement search => Driver.FindElement(By.XPath(name));
        IWebElement un => Driver.FindElement(By.XPath(uname));
        IWebElement res_dup => Driver.FindElement(By.XPath(code));
        // IWebElement res_input => Driver.FindElement(By.XPath(input));


        public void Get_Counterparty()
        {

            test = ExtReport.CreateTest("TC_031_GetCounterparty").Info("GetCounterparty Test has Started");

            string inputD = InputExcelAPI.GetCellData("WebAPI", 25, 3);
            /*string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 38, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 38, 4);*/

            SqlCommand cmd = new SqlCommand("select * from Counterparty where Code = @Code", con);
            cmd.Parameters.AddWithValue("@Code", inputD);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string cname;
            string gcdId;
            cname = Convert.ToString(reader[2]);
            gcdId = Convert.ToString(reader[3]);

            reader.Close();
            con.Close();

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));*/

            //Thread.Sleep(2000);

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(url)));

            url_dup.Click();

            test.Log(Status.Info, "GetCounterparty selected");
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(name)));

            search.SendKeys(inputD);
            test.Log(Status.Info, "Name : "+inputD+ " entered");

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(uname)));

            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(Try)));

            button_dup.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
            //Thread.Sleep(3000);
            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_dup);
            action.Perform();

            string resBody = body_dup.Text;
            string resCode = res_dup.Text;

            //test.Log(Status.Info, resBody);
            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GetCounterparty Response is " + resCode);

                if (resBody.Contains(cname)&&resBody.Contains(gcdId))
                {
                    test.Log(Status.Pass, "GetCounterparty Response body contain COunterparty name " + cname + " & GCDID : " + gcdId);
                    test.Log(Status.Pass, "Test Passed!");
                }
                else
                {
                    test.Log(Status.Fail, "GetCounterparty Response is failed!! not contain " + cname);
                    Assert.Fail("GetCounterparty Response is Failed!");
                }
            }
            else
            {
                test.Log(Status.Fail, "GetCounterparty Response is " + resCode);
                Assert.Fail("GetCounterparty Response is " + resCode);
            }

            url_dup.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);

        }
    }
}
