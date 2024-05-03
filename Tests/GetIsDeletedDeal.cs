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
    class GetIsDeletedDeal
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest ExtTest;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string
        DeletedDeal_url = "//li[@id='Deal_Deal_GetIsDeletedDeal']/div/h3/span/a[@class='toggleOperation']",
        DeletedDeal_Try = "//*[@id='Deal_Deal_GetIsDeletedDeal_content']/form/div[2]/input",
        Res_code = "//*[@id='Deal_Deal_GetIsDeletedDeal_content']/div[2]/div[4]/pre",
        DeletedDeal_body = "//*[@id='Deal_Deal_GetIsDeletedDeal_content']/div[2]/div[3]/pre/code",
        lastdealchangeId = "//div[@id='Deal_Deal_GetIsDeletedDeal_content']/form/table/tbody/tr/td/input[@name='_lastdealchangeId']",
        userName = "//div[@id='Deal_Deal_GetIsDeletedDeal_content']/form/table/tbody/tr/td/input[@name='Username']";

        public GetIsDeletedDeal(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_DeletedDeal => Driver.FindElement(By.XPath(DeletedDeal_url));
        IWebElement button_DeletedDeal => Driver.FindElement(By.XPath(DeletedDeal_Try));
        IWebElement body_DeletedDeal => Driver.FindElement(By.XPath(DeletedDeal_body));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
        IWebElement DeletedID => Driver.FindElement(By.XPath(lastdealchangeId));
        IWebElement UN => Driver.FindElement(By.XPath(userName));

        public void Deleted_Deal()
        {
            ExtTest = ExtReport.CreateTest("TS_062_Deal_GetIsDeletedDeal()").Info("Test Started");

            string value = InputExcelAPI.GetCellData("DatabaseTopic", 99, 3);
            string value1 = InputExcelAPI.GetCellData("DatabaseTopic", 99, 4);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(DeletedDeal_url)));

            url_DeletedDeal.Click();
            ExtTest.Log(Status.Info, "GetIsDeletedDeal API Call selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lastdealchangeId)));

            DeletedID.SendKeys("21261");
            ExtTest.Log(Status.Info, "DealID: 21261 entered");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            UN.SendKeys(username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(DeletedDeal_Try)));

            button_DeletedDeal.Click();
            ExtTest.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(DeletedDeal_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_DeletedDeal);
            action.Perform();

            string resBody = body_DeletedDeal.Text;
            string resCode = res_Code.Text;


            ExtTest.Log(Status.Info, "Verifying GetIsDeletedDeal Values....");

            if (resCode == "200")
            {
                ExtTest.Log(Status.Pass, "GetIsDeletedDeal Response is " + resCode);

                if (resBody.Contains(value) && resBody.Contains(value1))
                {
                    ExtTest.Log(Status.Pass, "GetIsDeletedDeal Response body contains " + value + " " + value1);
                    ExtTest.Log(Status.Pass, "Test case is Pass");
                }
                else
                {
                    ExtTest.Log(Status.Fail, "GetIsDeletedDeal Response is fail");
                    Assert.Fail("GetIsDeletedDeal Response is "+ resBody);
                }
            }
            else
            {
                ExtTest.Log(Status.Fail, "GetIsDeletedDeal Response is " + resCode);
                Assert.Fail("GetIsDeletedDeal Response is "+ resCode);
            }

            url_DeletedDeal.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }

    }

}