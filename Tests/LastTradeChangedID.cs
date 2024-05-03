using AventStack.ExtentReports;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SwaggerWebAPI.Libs;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace SwaggerWebAPI
{
    class LastTradeChangedID
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='Trade_Trade_GetlastchangeTradeChangeId']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='Trade_Trade_GetlastchangeTradeChangeId_content']/form/div[2]/input",
                         lattise_code = "//*[@id='Trade_Trade_GetlastchangeTradeChangeId_content']/div[2]/div[4]/pre",
                         userName = "//div[@id='Trade_Trade_GetlastchangeTradeChangeId_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='Trade_Trade_GetlastchangeTradeChangeId_content']/div[2]/div[3]/pre/code";


        public LastTradeChangedID(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_lattise => Driver.FindElement(By.XPath(lattise_url));
        IWebElement button_lattise => Driver.FindElement(By.XPath(lattise_Try));
        IWebElement body_lattise => Driver.FindElement(By.XPath(lattise_body));
        IWebElement un => Driver.FindElement(By.XPath(userName));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetLastTradeChangedID()
        {

            test = ExtReport.CreateTest("TS_089_GetLastTradeChangedID").Info("Test Started");

            //string val = ValidationExcelAPI.GetCellData("DatabaseTopic", 153, 3);

             SqlCommand cmd = new SqlCommand("select max(Id) from TradeChange", con);

             con.Open();
             SqlDataReader reader = cmd.ExecuteReader();

             reader.Read();

            string lastId;
            //string code;
            lastId = Convert.ToString(reader[0]);
             //code = Convert.ToString(reader[1]);

             reader.Close();
             con.Close();

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
            url_lattise.Click();
            test.Log(Status.Info, "LastTradeChangedID selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            string input = string.Format(username);
            un.SendKeys(username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_Try)));
            button_lattise.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lattise_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_lattise);
            action.Perform();

            string resBody = body_lattise.Text;
            string resCode = res_Code.Text;


            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "LastTradeChangedID Response is " + resCode);

                if (resBody.Contains("-1"))
                {
                    test.Log(Status.Pass, "There are no deals stuck at blue color state in the Lattise app");
                    test.Log(Status.Pass, "Test is Pass");
                }

                else if (resBody.Contains(lastId))
                {
                    test.Log(Status.Warning, "Deals are stuck at blue color state!!");
                    Assert.Warn("LastTradeChangedID Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "LastTradeChangedID Response is " + resCode);
                Assert.Fail("LastTradeChangedID Response code is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
