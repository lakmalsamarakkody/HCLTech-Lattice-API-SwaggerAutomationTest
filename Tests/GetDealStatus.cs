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
    class GetDealStatus
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest ExtTest;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string
        DealStatus_url = "//li[@id='Deal_Deal_GetDealStatus']/div/h3/span/a[@class='toggleOperation']",
        DealStatus_Try = "//*[@id='Deal_Deal_GetDealStatus_content']/form/div[2]/input",
        Res_code = "//*[@id='Deal_Deal_GetDealStatus_content']/div[2]/div[4]/pre",
        DealStatus_body = "//*[@id='Deal_Deal_GetDealStatus_content']/div[2]/div[3]/pre/code",
        userName = "//*[@id='Deal_Deal_GetDealStatus_content']/form/table/tbody/tr/td/input[@name='Username']";



        public GetDealStatus(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_DealStatus => Driver.FindElement(By.XPath(DealStatus_url));
        IWebElement button_DealStatus => Driver.FindElement(By.XPath(DealStatus_Try));
        IWebElement body_DealStatus => Driver.FindElement(By.XPath(DealStatus_body));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
        IWebElement UN => Driver.FindElement(By.XPath(userName));

        public void Deal_Status()
        {
            ExtTest = ExtReport.CreateTest("TS_063_Deal_GetDealStatus()").Info("Test Started");
            //   string value = ExcelAPI.GetCellData("DatabaseTopic", 101, 3);
            //   string value1 = ExcelAPI.GetCellData("DatabaseTopic", 101, 4);
            SqlCommand cmd = new SqlCommand("select  distinct(DealChangeID) from TradeChange where Status IN (1,2,3) order by DealChangeID desc", con);
            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            string Id;
          //  string dealId;

            reader.Read();

            Id = Convert.ToString(reader[0]);
         //   dealId = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(DealStatus_url)));

            url_DealStatus.Click();
            ExtTest.Log(Status.Info, "GetDealStatus API Call selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            string input = string.Format(username);
            UN.SendKeys(username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(DealStatus_Try)));

            button_DealStatus.Click();
            ExtTest.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(DealStatus_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_DealStatus);
            action.Perform();

            string resBody = body_DealStatus.Text;
            string resCode = res_Code.Text;


            ExtTest.Log(Status.Info, "Verifying GetDealStatus Values....");

            if (resCode == "200")
            {
                ExtTest.Log(Status.Pass, "GetDealStatus Response is " + resCode);

                if (resBody.Contains(Id))
                {
                    ExtTest.Log(Status.Pass, "GetDealStatus Response body contains " + Id);
                    ExtTest.Log(Status.Pass, "Test case is Pass");
                }
                else
                {
                    ExtTest.Log(Status.Fail, "GetDealStatus Response is fail");
                    Assert.Fail("GetDealStatus Response is "+ Id);
                }
            }
            else
            {
                ExtTest.Log(Status.Fail, "GetDealStatus Response is " + resCode);
                Assert.Fail("GetDealStatus Response is "+ resCode);
            }

            url_DealStatus.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }

    }
}