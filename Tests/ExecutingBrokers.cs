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
    class ExecutingBrokers
    {

        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();

        static string Broker_url = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetExecutingBrokers']/div[1]/h3/span[1]/a",
                         Broker_Try = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetExecutingBrokers_content']/form/div[2]/input",
                         Res_code = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetExecutingBrokers_content']/div[2]/div[4]/pre",
                         Username = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetExecutingBrokers_content']/form/table/tbody/tr/td/input[@name='Username']",
                        Broker_body = "//*[@id='CounterPartyTraderBroker_CounterPartyTraderBroker_GetExecutingBrokers_content']/div[2]/div[3]/pre/code";

        public ExecutingBrokers(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_Broker => Driver.FindElement(By.XPath(Broker_url));
        IWebElement button_Broker => Driver.FindElement(By.XPath(Broker_Try));
        IWebElement body_Broker => Driver.FindElement(By.XPath(Broker_body));
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));


        public void GetExcBroker()
        {

            test = ExtReport.CreateTest("TS_006_ExcutingBrokers").Info("Test Started");

            //string val1= ValidationExcelAPI.GetCellData("CounterPartyTradeBroker", 1, 2);
            string val2= ValidationExcelAPI.GetCellData("CounterPartyTradeBroker", 1, 3);

            SqlCommand cmd = new SqlCommand("select * from BrokerChange where Id=(select max(Id) from BrokerChange)", con);
           // cmd.Parameters.AddWithValue("@Name", val2);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string name, postingID;
           
           // string code;
         
            postingID = Convert.ToString(reader[1]);
            name = Convert.ToString(reader[2]);

            reader.Close();
            con.Close();
            // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Broker_url)));
            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(Broker_url))).Click();

            //url_Broker.Click();

           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
            test.Log(Status.Info, "Executing Broker selected");

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(Broker_Try)));

            button_Broker.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(Broker_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_Broker);
            action.Perform();

            string resBody = body_Broker.Text;
            string resCode = res_Code.Text;

            
            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "Executing Broker Response is " + resCode);

                if (resBody.Contains(postingID) && resBody.Contains(name))
                {
                    test.Log(Status.Pass, "Executing Broker Response body contains posting ID " + postingID+ " & Broker "+name);
                    test.Log(Status.Pass, "Executing Broker is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "Executing Broker Response is failed! not contain "+ postingID + " & "+name);
                    Assert.Fail("Executing Broker Response not containing "+ name);

                }
            }
            else
            {
                test.Log(Status.Fail, "Executing Broker Response is " + resCode);
                Assert.Fail("Executing Broker Response is "+ resCode);

            }
            url_Broker.Click();
           Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }

    }
}
