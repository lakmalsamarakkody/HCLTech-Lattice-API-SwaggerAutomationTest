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
    class GetClients
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();
        //*[@id="CounterPartyTraderBroker_CounterPartyTraderBroker_GetClients"]/div[1]/h3/span[1]/a
        //*[@id="CounterPartyTraderBroker_CounterPartyTraderBroker_GetClients"]/div[1]/h3/span[1]/a

        static string Trade_url = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetClients']/div[1]/h3/span[1]/a",
                         Trade_Try = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetClients_content']/form/div[2]/input",
                         Res_code = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetClients_content']/div[2]/div[4]/pre",
                        Trade_body = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetClients_content']/div[2]/div[3]/pre/code",
                        Username = "//div[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetClients_content']/form/table/tbody/tr/td/input[@name='Username']";

        public GetClients(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_Trade => Driver.FindElement(By.XPath(Trade_url));
        IWebElement button_Trade => Driver.FindElement(By.XPath(Trade_Try));
        IWebElement body_Trade => Driver.FindElement(By.XPath(Trade_body));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
        IWebElement un => Driver.FindElement(By.XPath(Username));


        public void GetTradeBroker()
        {

            test = ExtReport.CreateTest("TS_005_GetClients").Info("Test Started");

           // string value = ValidationExcelAPI.GetCellData("CounterPartyTradeBroker", 1, 0);

            SqlCommand cmd = new SqlCommand("select TOP(1) p.ID ,t.TraderCode,p.CounterPartyID,p.TraderID from CounterPartyTraderBrokerCode as p join CounterPartyTrader as t on  p.TraderID = t.ID where p.IsActive = 1 order by p.TraderID DESC", con);
           // cmd.Parameters.AddWithValue("@Name", value);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string name;
            //string code;
             //code = Convert.ToString(reader[1]);
            name = Convert.ToString(reader[1]).Trim();

            reader.Close();
            con.Close();

            // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Trade_url)));
            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(Trade_url))).Click();

           // url_Trade.Click();

            test.Log(Status.Info, "TradeBroker selected");
            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            un.SendKeys(username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Trade_Try)));
            button_Trade.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");
           
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Trade_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_Trade);
            action.Perform();

           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(4);

            string resBody = body_Trade.Text;
            string resCode = res_Code.Text;

            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "Trade Broker Response is " + resCode);

                if (resBody.Contains(name))
                {
                    test.Log(Status.Pass, "Trade Broker Response body contains Trader name " + name);
                    test.Log(Status.Pass, "CounterpartyTraderBroker getclient API call pass");
                }
                else
                {
                    test.Log(Status.Fail, "Trade Broker Response is fail");
                    Assert.Fail("GetCLient Response is not contain");
                }
            }
            else
            {
                test.Log(Status.Fail, "CcyPair Response is " + resCode);
                Assert.Fail("GetCLient Response is ", resCode);

            }
            url_Trade.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }
    }
}
