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
    class GetLatestDealChange
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest ExtTest;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string
        LatestDealchange_url = "//li[@id='Deal_Deal_GetLatestDealChange']/div/h3/span/a[@class='toggleOperation']",
        LatestDealchange_Try = "//*[@id='Deal_Deal_GetLatestDealChange_content']/form/div[2]/input",
        Res_code = "//*[@id='Deal_Deal_GetLatestDealChange_content']/div[2]/div[4]/pre",
        LatestDealchange_body = "//*[@id='Deal_Deal_GetLatestDealChange_content']/div[2]/div[3]/pre/code",
        userName = "//div[@id='Deal_Deal_GetLatestDealChange_content']/form/table/tbody/tr/td/input[@name='Username']";

        public GetLatestDealChange(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_LatestDealchange => Driver.FindElement(By.XPath(LatestDealchange_url));
        IWebElement button_LatestDealchange => Driver.FindElement(By.XPath(LatestDealchange_Try));
        IWebElement body_LatestDealchange => Driver.FindElement(By.XPath(LatestDealchange_body));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
        IWebElement UN => Driver.FindElement(By.XPath(userName));

        public void LatestDeal_Change()
        {
            ExtTest = ExtReport.CreateTest("TS_076_GetLatestDealChange()").Info("Test Started");
            //  string value = ExcelAPI.GetCellData("DatabaseTopic", 103, 3);

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

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(LatestDealchange_url)));

            url_LatestDealchange.Click();
            ExtTest.Log(Status.Info, "GetLatestDealChange API Call selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            string input = string.Format(username);
            UN.SendKeys(username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(LatestDealchange_Try)));

            button_LatestDealchange.Click();
            ExtTest.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(LatestDealchange_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_LatestDealchange);
            action.Perform();

            string resBody = body_LatestDealchange.Text;
            string resCode = res_Code.Text;

            ExtTest.Log(Status.Info, "Verifying GetLatestDealChange Values....");

            if (resCode == "200")
            {
                ExtTest.Log(Status.Pass, "GetLatestDealChange Response is " + resCode);

                if (resBody.Contains(lastId))
                {
                    ExtTest.Log(Status.Pass, "GetLatestDealChange Response body contains " + lastId);
                    ExtTest.Log(Status.Pass, "Test case is Pass");
                }
                else
                {
                    ExtTest.Log(Status.Fail, "GetLatestDealChange Response is fail");
                    Assert.Fail("GetLatestDealChange Response is "+ resBody);
                }
            }
            else
            {
                ExtTest.Log(Status.Fail, "GetLatestDealChange Response is " + resCode);
                Assert.Fail("GetLatestDealChange Response is code "+ resCode);
            }

            url_LatestDealchange.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }

    }
}