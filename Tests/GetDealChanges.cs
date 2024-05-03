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
    class GetDealChanges
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest ExtTest;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string
        DealChanges_url = "//li[@id='Deal_Deal_GetDealChanges']/div/h3/span/a[@class='toggleOperation']",
        DealChanges_Try = "//*[@id='Deal_Deal_GetDealChanges_content']/form/div[2]/input",
        Res_code = "//*[@id='Deal_Deal_GetDealChanges_content']/div[2]/div[4]/pre",
        DealChanges_body = "//*[@id='Deal_Deal_GetDealChanges_content']/div[2]/div[3]/pre/code",
        LastchangeId = "//div[@id='Deal_Deal_GetDealChanges_content']/form/table/tbody/tr/td/input[@name='_lastChangeId']",
        userName = "//div[@id='Deal_Deal_GetDealChanges_content']/form/table/tbody/tr/td/input[@name='Username']";




        public GetDealChanges(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_DealChanges => Driver.FindElement(By.XPath(DealChanges_url));
        IWebElement button_DealChanges => Driver.FindElement(By.XPath(DealChanges_Try));
        IWebElement body_DealChanges => Driver.FindElement(By.XPath(DealChanges_body));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
        IWebElement Lastchange_DealId => Driver.FindElement(By.XPath(LastchangeId));
        IWebElement UN => Driver.FindElement(By.XPath(userName));

        public void Get_DealChanges()
        {
            ExtTest = ExtReport.CreateTest("TS_066_Deal_GetDealChanges()").Info("Test Started");

           SqlCommand cmd = new SqlCommand("select TOP(1)* FROM DealChange where id in (SELECT TOP(2) Id FROM DealChange order by id desc)order by id asc", con);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            string lastId;
            string dealId;

            reader.Read();

            lastId = Convert.ToString(reader[0]);
            dealId = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();

            //  string value = ExcelAPI.GetCellData("DatabaseTopic", 105, 3);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(DealChanges_url)));

            url_DealChanges.Click();
            ExtTest.Log(Status.Info, "GetDealChanges API Call selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(LastchangeId)));
            Lastchange_DealId.SendKeys(lastId);
          //  ExtTest.Log(Status.Info, "DealID: 21194 entered");
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            UN.SendKeys(username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(DealChanges_url)));

            button_DealChanges.Click();
            ExtTest.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(DealChanges_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_DealChanges);
            action.Perform();

            string resBody = body_DealChanges.Text;
            string resCode = res_Code.Text;


            ExtTest.Log(Status.Info, "Verifying GetDealCahnges Values....");

            if (resCode == "200")
            {
                ExtTest.Log(Status.Pass, "GetDealCahnges Response is " + resCode);

                if (resBody.Contains(dealId))
                {
                    ExtTest.Log(Status.Pass, "GetDealCahnges Response body contains " + dealId);
                    ExtTest.Log(Status.Pass, "Test case is Pass");
                }
                else
                {
                    ExtTest.Log(Status.Fail, "GetDealCahnges Response is fail");
                    Assert.Fail("GetDealCahnges Response is "+ resBody);
                }
            }
            else
            {
                ExtTest.Log(Status.Fail, "GetDealCahnges Response is " + resCode);
                Assert.Fail("GetDealCahnges Response is "+ resCode);
            }

            url_DealChanges.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }

    }
}

