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
    class CcyPairs
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest ExtTest;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string
        CcyPair_url = "//li[@id='CcyPairs_CcyPairs_GetCcyPairs']/div/h3/span/a[@class='toggleOperation']",
        Ccy_Try = ".//*[@id='CcyPairs_CcyPairs_GetCcyPairs_content']/form/div[2]/input",
        Res_code = "//*[@id='CcyPairs_CcyPairs_GetCcyPairs_content']/div[2]/div[4]/pre",
        Ccy_body = "//*[@id='CcyPairs_CcyPairs_GetCcyPairs_content']/div[2]/div[3]/pre/code",
        userName = "/html/body/div[3]/div[2]/ul/li[1]/ul/li/ul/li/div[2]/form/table/tbody/tr/td[2]/input";

        public CcyPairs(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_CcyPair => Driver.FindElement(By.XPath(CcyPair_url));
        IWebElement button_CcyPair => Driver.FindElement(By.XPath(Ccy_Try));
        IWebElement body_CcyPair => Driver.FindElement(By.XPath(Ccy_body));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
        IWebElement UN => Driver.FindElement(By.XPath(userName));

        public void GetCcyPairs()
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

            SqlCommand cmd = new SqlCommand("select * from CcyPair where Id = (select max(Id) from CcyPair)", con);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            string code;

            reader.Read();

            
            code = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();




            ExtTest = ExtReport.CreateTest("TS_001_CcyPairs_GetCcyPairs").Info("Test Started");
            string value = ValidationExcelAPI.GetCellData("CcyPair", 1, 0);
            url_CcyPair.Click();
            ExtTest.Log(Status.Info, "CcyPair API Call selected");
            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(userName)));

            UN.SendKeys(username);
            button_CcyPair.Click();
            ExtTest.Log(Status.Info, "Try it Now Button Clicked");
            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(Ccy_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_CcyPair);
            action.Perform();

            string resBody = body_CcyPair.Text;
            string resCode = res_Code.Text;


            ExtTest.Log(Status.Info, "Verifying CcyPair Values....");

            if (resCode == "200")
            {
                ExtTest.Log(Status.Pass, "CcyPair Response is " + resCode);

                if (resBody.Contains(code))
                {
                    ExtTest.Log(Status.Pass, "GetNewlyaddedCCyPair Response body contains last added CcyPair "+code);
                    ExtTest.Log(Status.Pass, "Test case is Pass");
                }
                else
                {
                    ExtTest.Log(Status.Fail, "CcyPair Response is fail");
                    Assert.Fail("CcyPair Response is ", value);
                }
            }
            else
            {
                ExtTest.Log(Status.Fail, "CcyPair Response is " + resCode);
                Assert.Fail("CcyPair Response is ", resCode);
            }

            url_CcyPair.Click();
           Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }
    }
}
