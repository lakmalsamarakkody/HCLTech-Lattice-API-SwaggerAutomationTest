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
namespace SwaggerWebAPI.Tests
{
    class GetDeal
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest ExtTest;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string
        Deal_url = "//li[@id='Deal_Deal_GetDeal']/div/h3/span/a[@class='toggleOperation']",
        Deal_Try = "//*[@id='Deal_Deal_GetDeal_content']/form/div[2]/input",
        Res_code = "//*[@id='Deal_Deal_GetDeal_content']/div[2]/div[4]/pre",
        Deal_body = "//*[@id='Deal_Deal_GetDeal_content']/div[2]/div[3]/pre/code",
        DealID = "//div[@id='Deal_Deal_GetDeal_content']/form/table/tbody/tr/td/input[@name='dealid']",
        userName = "//div[@id='Deal_Deal_GetDeal_content']/form/table/tbody/tr/td/input[@name='Username']";

        public GetDeal(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_Deal => Driver.FindElement(By.XPath(Deal_url));
        IWebElement button_Deal => Driver.FindElement(By.XPath(Deal_Try));
        IWebElement body_Deal => Driver.FindElement(By.XPath(Deal_body));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
        IWebElement ID_Deal => Driver.FindElement(By.XPath(DealID));
        IWebElement UN => Driver.FindElement(By.XPath(userName));

        public void Get_Deal()
        {
            ExtTest = ExtReport.CreateTest("TS_074_GetDeal").Info("Test Started");

            string value = InputExcelAPI.GetCellData("WebAPI", 44, 3);

            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 117, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 117, 4);
            string val3 = ValidationExcelAPI.GetCellData("DatabaseTopic", 117, 5);
            string val4 = ValidationExcelAPI.GetCellData("DatabaseTopic", 117, 6);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(Deal_url)));
            url_Deal.Click();
            ExtTest.Log(Status.Info, "GetDeal API Call selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(DealID)));
            ID_Deal.SendKeys(value);
            ExtTest.Log(Status.Info, "DealID: "+value+ " entered");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            string input = string.Format(username);
            UN.SendKeys(username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(Deal_Try)));
            button_Deal.Click();
            ExtTest.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Deal_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_Deal);
            action.Perform();

            string resBody = body_Deal.Text;
            string resCode = res_Code.Text;


            ExtTest.Log(Status.Info, "Verifying GetDeal Values....");

            if (resCode == "200")
            {
                ExtTest.Log(Status.Pass, "GetDeal Response is " + resCode);

                if (resBody.Contains(val1) && resBody.Contains(val2) && resBody.Contains(val3) && resBody.Contains(val4))
                {
                    ExtTest.Log(Status.Pass, "GetDeal Response contains DealID " + val1 + " , DealChangedID " + val2 + " & TradeIDs " + val3 + ", " + val4);
                    ExtTest.Log(Status.Pass, "Test case is Pass");
                }
                else
                {
                    ExtTest.Log(Status.Fail, "GetDeal Response is fail");
                    Assert.Fail("GetDeal Response is ", value);
                }
            }
            else
            {
                ExtTest.Log(Status.Fail, "GetDeal Response is " + resCode);
                Assert.Fail("GetDeal Response is ", resCode);
            }

            url_Deal.Click();
          //  Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }
    
        }
}

