using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using SwaggerWebAPI.Libs;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Threading;
namespace SwaggerWebAPI.Tests
{
    class PrimeBrokerIds
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest ExtTest;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string
        PBId_url = "//li[@id='DatabaseTopic_DatabaseTopic_GetPrimeBrokerIds']/div/h3/span/a[@class='toggleOperation']",
        PBId_Try =   "//*[@id='DatabaseTopic_DatabaseTopic_GetPrimeBrokerIds_content']/form/div[2]/input",
        Res_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetPrimeBrokerIds_content']/div[2]/div[4]/pre",
        PBId_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetPrimeBrokerIds_content']/div[2]/div[3]/pre/code",
        userName = "//div[@id='DatabaseTopic_DatabaseTopic_GetPrimeBrokerIds_content']/form/table/tbody/tr/td/input[@name='Username']";

        public PrimeBrokerIds(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_PrimeBrokerId => Driver.FindElement(By.XPath(PBId_url));
        IWebElement button_PrimeBrokerId => Driver.FindElement(By.XPath(PBId_Try));
        IWebElement body_PrimeBrokerId => Driver.FindElement(By.XPath(PBId_body));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
        IWebElement UN => Driver.FindElement(By.XPath(userName));

        public void GetPrimeBrokerIds()
        {

            //SqlCommand cmd = new SqlCommand("Update CcyPair set AddedBy = @AddedBy where Id = @Id", con);
            //cmd.Parameters.AddWithValue("@AddedBy", "Lakmal");
            //cmd.Parameters.AddWithValue("@Id", 1);
            //con.Open();
            //int result = cmd.ExecuteNonQuery();
            //con.Close();
            //if (result > 0)
            //{
            //    Console.WriteLine("CcyPair table has been updated");
            //}

            ExtTest = ExtReport.CreateTest("TS_010_GetPrimeBrokerIds()").Info("Test Started");
            string value = ValidationExcelAPI.GetCellData("DatabaseTopic", 6, 3);

           // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(PBId_url)));
            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(PBId_url)));

           url_PrimeBrokerId.Click();
            ExtTest.Log(Status.Info, "GetPrimeBrokerIds API Call selected");
            // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
           //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

            string input = string.Format(username);
            UN.SendKeys(input);

            button_PrimeBrokerId.Click();
            ExtTest.Log(Status.Info, "Try it Now Button Clicked");
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(PBId_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_PrimeBrokerId);
            action.Perform();

            string resBody = body_PrimeBrokerId.Text;
            string resCode = res_Code.Text;

            ExtTest.Log(Status.Info, "Verifying GetPrimeBrokerIds Values....");

            if (resCode == "200")
            {
                ExtTest.Log(Status.Pass, "GetPrimeBrokerIds Response is " + resCode);

                if (resBody.Contains(value))
                {
                    ExtTest.Log(Status.Pass, "GetPrimeBrokerIds Response body contains " + value);
                    ExtTest.Log(Status.Pass, "Test case is Pass");
                }
                else
                {
                    ExtTest.Log(Status.Fail, "GetPrimeBrokerIds Response is fail");
                    Assert.Fail("GetPrimeBrokerIds Response is ", value);
                }
            }
            else
            {
                ExtTest.Log(Status.Fail, "GetPrimeBrokerIds Response is " + resCode);
                Assert.Fail("GetPrimeBrokerIds Response is "+resCode);
            }

            url_PrimeBrokerId.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }
    }
}

