using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SwaggerWebAPI.Libs;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Threading;

namespace SwaggerWebAPI
{
    class GetLatestDealChangeId
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest ExtTest;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string
        LatestDealchangeid_url = "//*[@id='Deal_Deal_GetLatestDealChangeId']/div[1]/h3/span[2]/a",
        LatestDealchangeid_Try = "//*[@id='Deal_Deal_GetLatestDealChangeId_content']/form/div[2]/input",
        Res_code = "//*[@id='Deal_Deal_GetLatestDealChangeId_content']/div[2]/div[4]/pre",
        LatestDealchangeid_body = "//*[@id='Deal_Deal_GetLatestDealChangeId_content']/div[2]/div[3]/pre/code",
        userName = "//div[@id='Deal_Deal_GetLatestDealChangeId_content']/form/table/tbody/tr/td/input[@name='Username']";
        
        public GetLatestDealChangeId(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_LatestDealchangeid => Driver.FindElement(By.XPath(LatestDealchangeid_url));
        IWebElement button_LatestDealchangeid => Driver.FindElement(By.XPath(LatestDealchangeid_Try));
        IWebElement body_LatestDealchangeid => Driver.FindElement(By.XPath(LatestDealchangeid_body));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
        IWebElement UN => Driver.FindElement(By.XPath(userName));

        public void LatestDeal_ChangeId()
        {
            ExtTest = ExtReport.CreateTest("TS_077_GetLatestDealChangeId()").Info("Test Started");
            
            SqlCommand cmd = new SqlCommand("select* from DealChange where Id = (select max(Id) from DealChange)", con);
            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            string lastId;
            string dealId;

            reader.Read();

            lastId = Convert.ToString(reader[0]);
            dealId = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(LatestDealchangeid_url)));

            url_LatestDealchangeid.Click();
            ExtTest.Log(Status.Info, "GetLatestDealChangeId API Call selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            string input = string.Format(username);
            UN.SendKeys(username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(LatestDealchangeid_Try)));

            button_LatestDealchangeid.Click();
            ExtTest.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(LatestDealchangeid_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_LatestDealchangeid);
            action.Perform();

            string resBody = body_LatestDealchangeid.Text;
            string resCode = res_Code.Text;


            ExtTest.Log(Status.Info, "Verifying GetLatestDealChangeId Values....");

            if (resCode == "200")
            {
                ExtTest.Log(Status.Pass, "GetLatestDealChangeId Response is " + resCode);

                if (resBody.Contains(lastId))
                {
                    ExtTest.Log(Status.Pass, "GetLatestDealChangeId Response body contains " + lastId + "&dealId" + dealId);
                    ExtTest.Log(Status.Pass, "Test case is Pass");
                }
                else
                {
                    ExtTest.Log(Status.Fail, "GetLatestDealChangeId Response is fail");
                    Assert.Fail("GetLatestDealChangeId Response is "+ lastId);
                }
            }
            else
            {
                ExtTest.Log(Status.Fail, "GetLatestDealChangeId Response is " + resCode);
                Assert.Fail("GetLatestDealChangeId Response code is "+ resCode);
            }

            url_LatestDealchangeid.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }

    }
}