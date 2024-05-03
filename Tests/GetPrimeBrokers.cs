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
    class GetPrimeBrokers
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetPrimeBrokers']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetPrimeBrokers_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetPrimeBrokers_content']/div[2]/div[4]/pre",
                         userName = "//*[@id='DatabaseTopic_DatabaseTopic_GetPrimeBrokers_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetPrimeBrokers_content']/div[2]/div[3]/pre/code";


        public GetPrimeBrokers(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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


        public void Get_PrimeBrokers()
        {

            test = ExtReport.CreateTest("TS_056_GetPrimeBrokers").Info("Test Started");
/*
            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 85, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 85, 4);*/

            SqlCommand cmd = new SqlCommand("select top(1) p.Id,c.Code from PrimeBrokerCounterparty as p join CounterParty as c on p.CounterpartyID = c.Id order by p.Id desc", con);

            //cmd.Parameters.AddWithValue("@username", User.GetUsername());

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string code, id;

            code = Convert.ToString(reader[1]).Trim();
            id = Convert.ToString(reader[0]);

            reader.Close();
            con.Close();
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "GetPrimeBrokers selected");

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


            test.Log(Status.Info, "Verifying Value....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GetPrimeBrokers Response is " + resCode);

                if (resBody.Contains(id)&& resBody.Contains(code))
                {
                    test.Log(Status.Pass, "GetPrimeBrokers Response body contains PrimebrokerCounterparty ID:  " + id + " & Code " + code);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetPrimeBrokers Response is Failed");
                    Assert.Fail("GetPrimeBrokers Response code is Faile!!");
                }
            }
            else
            {
                test.Log(Status.Fail, "GetPrimeBrokers Response is " + resCode);
                Assert.Fail("GetPrimeBrokers Response code is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
