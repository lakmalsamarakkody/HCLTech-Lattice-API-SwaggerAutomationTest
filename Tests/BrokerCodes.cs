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
using System.Threading.Tasks;

namespace SwaggerWebAPI
{
    class BrokerCodes
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();

        static string BCode_url = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetBrokerCodes']/div[1]/h3/span[1]/a",
                         BCode_Try = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetBrokerCodes_content']/form/div[2]/input",
                         Res_code = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetBrokerCodes_content']/div[2]/div[4]/pre",
                         Username = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetBrokerCodes_content']/form/table/tbody/tr/td/input[@name='Username']",
                        BCode_body = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetBrokerCodes_content']/div[2]/div[3]/pre/code";

        public BrokerCodes(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_BCode => Driver.FindElement(By.XPath(BCode_url));
        IWebElement button_BCode => Driver.FindElement(By.XPath(BCode_Try));
        IWebElement body_BCode => Driver.FindElement(By.XPath(BCode_body));
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));


        public void GetBrokerCodes()
        {

            test = ExtReport.CreateTest("TC_006_BrokerCodes").Info("Test Started");

           // string val1 = ValidationExcelAPI.GetCellData("CounterPartyTradeBroker", 1, 3);
            string val = ValidationExcelAPI.GetCellData("CounterPartyTradeBroker", 1, 6);

            SqlCommand cmd = new SqlCommand("select * from BrokerCode where Code = @Code", con);
            cmd.Parameters.AddWithValue("@Code", val);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string id;
            string code;
            code = Convert.ToString(reader[1]);
            id = Convert.ToString(reader[0]);

            reader.Close();
            con.Close();

            //WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(BCode_url)));
            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BCode_url))).Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
//            url_BCode.Click();

            test.Log(Status.Info, "BrokerCodes selected");
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
            // WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(Username)));
            
            string input = string.Format(username);
           // un.SendKeys(input);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username))).SendKeys(input);

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BCode_Try)));

            button_BCode.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BCode_body)));
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);

            Actions action = new Actions(Driver);
            action.MoveToElement(body_BCode);
            action.Perform();

            string resBody = body_BCode.Text;
            string resCode = res_Code.Text;

            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "BrokerCodes Response is " + resCode);

                if (resBody.Contains(id)&& resBody.Contains(code))
                {
                    test.Log(Status.Pass, "BrokerCodes Response body contains Broker ID " + id+ " & Counterparty "+code);
                    test.Log(Status.Pass, "Test 6 is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "BrokerCodes Response is failed!! not contain "+id+" & "+code);
                    Assert.Fail("BrokerCode Response not contain "+code );
                }
            }
            else
            {
                test.Log(Status.Fail, "BrokerCodes Response is " + resCode);
                Assert.Fail("BrokerCode Response is ", resCode);
            }

            url_BCode.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);

        }
    }
}
