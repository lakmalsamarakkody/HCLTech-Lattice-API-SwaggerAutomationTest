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
    class Getconnection
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest ExtTest;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string Connection_url = "//*[@id='Connection_Connection_GetConnection']/div[1]/h3/span[1]/a",
                      Conn_button = "//*[@id='Connection_Connection_GetConnection']/div[2]/form/div[2]/input",
                      Res_code = "//*[@id='Connection_Connection_GetConnection']/div[2]/div[2]/div[4]/pre",
                      Conn_body = "//*[@id='Connection_Connection_GetConnection']/div[2]/div[2]/div[3]/pre/code",
                      connection_username = "//*[@id='Connection_Connection_GetConnection_content']/form/table/tbody/tr/td/input[@name='Username']";

        public Getconnection(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement conn_url => Driver.FindElement(By.XPath(Connection_url));
        IWebElement button_Conn => Driver.FindElement(By.XPath(Conn_button));
        IWebElement body_Conn => Driver.FindElement(By.XPath(Conn_body));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
        IWebElement UN => Driver.FindElement(By.XPath(connection_username));

        public void GetConnection()
        {

            ExtTest = ExtReport.CreateTest("TS_002_Getconnection").Info("Test Started");
            string value = ValidationExcelAPI.GetCellData("Connection", 0, 0);

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(Connection_url))).Click();
            //conn_url.Click();

            ExtTest.Log(Status.Info, "CounterParty API Call selected");
           // new WebDriverWait(Driver, TimeSpan.FromSeconds(3)).Until(driver => Driver.FindElement(By.XPath(connection_username)));

           // WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(connection_username))).SendKeys(username);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(connection_username)));
            UN.SendKeys(username);
            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(Conn_button)));
            button_Conn.Click();
            ExtTest.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(Conn_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_Conn);
            action.Perform();

            string resBody = body_Conn.Text;
            string resCode = res_Code.Text;

            ExtTest.Log(Status.Info, "Verifying Connection Values....");

            if (resCode == "200")
            {
                ExtTest.Log(Status.Pass, "Connection Response is " + resCode);

                if (resBody.Contains(value))
                {
                    ExtTest.Log(Status.Pass, "Connection Response body contains " + value);
                    ExtTest.Log(Status.Pass, "Test case is Pass");
                }
                else
                {
                    ExtTest.Log(Status.Fail, "Connection Response body is fail " + value);
                    Assert.Fail("Connection Response body not contain this" + value);
                }
            }
            else
            {
                ExtTest.Log(Status.Fail, "Connection Response is " + resCode);
                Assert.Fail("Connection Response is ", resCode);
            }

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(Connection_url))).Click();

            //conn_url.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }


    }






}

