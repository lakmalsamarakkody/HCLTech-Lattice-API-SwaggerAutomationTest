using System;
using OpenQA.Selenium;
using System.Threading;
using AventStack.ExtentReports;
using AventStack.ExtentReports;
using NPOI.XSSF.UserModel;
using System.IO;
using OpenQA.Selenium.Interactions;
using SwaggerWebAPI.Libs;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using NUnit.Framework;
using System.Data.SqlClient;

namespace SwaggerWebAPI
{
    class LattiseCounterTradeDetails
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeCounterPartyTraderDetails']/div[1]/h3/span[1]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeCounterPartyTraderDetails_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeCounterPartyTraderDetails_content']/div[2]/div[4]/pre",
                         userName ="//div[@id='DatabaseTopic_DatabaseTopic_GetLatticeCounterPartyTraderDetails_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeCounterPartyTraderDetails_content']/div[2]/div[3]/pre/code";


        public LattiseCounterTradeDetails(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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


        public void GetLattiseCounterTradeDetails()
        {

            test = ExtReport.CreateTest("TS_033_GetLattiseCounterTradeDetails").Info("Test Started");

            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 40, 3);
           // string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 40, 4);

            SqlCommand cmd = new SqlCommand("select * from CounterPartyTrader where ID = @ID", con);
            cmd.Parameters.AddWithValue("@ID", val1);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string tname;
            string Id;
            tname = Convert.ToString(reader[2]);
            Id = Convert.ToString(reader[0]);

            reader.Close();
            con.Close();

            //string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "LattiseCounterTradeDetails selected");
            Thread.Sleep(2000);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));

            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_Try)));

            button_lattise.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lattise_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_lattise);
            action.Perform();

            string resBody = body_lattise.Text;
            string resCode = res_Code.Text;


            test.Log(Status.Info, "Verifying CodeResult Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "LattiseCounterTradeDetails Response is " + resCode);

                if (resBody.Contains(Id)&& resBody.Contains(tname))
                {
                    test.Log(Status.Pass, "LattiseCounterTradeDetails Response body contains ID " + Id + " & Name "+ tname);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "LattiseCounterTradeDetails Response is fail");
                    Assert.Fail("LattiseCounterTradeDetails Response is Failed!");
                }
            }
            else
            {
                test.Log(Status.Fail, "LattiseCounterTradeDetails Response is " + resCode);
                Assert.Fail("LattiseCounterTradeDetails Response is " + resCode);
            }

            url_lattise.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
