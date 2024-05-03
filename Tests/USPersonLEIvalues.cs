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
    class USPersonLEIvalues
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetGcdIdIsUSPersonLEIValues']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetGcdIdIsUSPersonLEIValues_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetGcdIdIsUSPersonLEIValues_content']/div[2]/div[4]/pre",
                         userName = "//*[@id='DatabaseTopic_DatabaseTopic_GetGcdIdIsUSPersonLEIValues_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetGcdIdIsUSPersonLEIValues_content']/div[2]/div[3]/pre/code";


        public USPersonLEIvalues(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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

        public void GetUSPersonLEIvalues()
        {

            test = ExtReport.CreateTest("TS_067_GetUSPersonLEIvalues").Info("Test Started");

            /*string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 103, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 103, 4);
            string val3 = ValidationExcelAPI.GetCellData("DatabaseTopic", 103, 5);*/

            SqlCommand cmd = new SqlCommand("select * from Counterparty where ID = (select max(ID) from Counterparty)", con);

            //cmd.Parameters.AddWithValue("@Code", value);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string cname;
            string usperson, id;

            cname = Convert.ToString(reader[2]).Trim();
            id = Convert.ToString(reader[0]);
            usperson = Convert.ToString(reader[8]).Trim();

            reader.Close();
            con.Close();

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "GetUSPersonLEIvalues selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
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
                test.Log(Status.Pass, "GetUSPersonLEIvalues Response is " + resCode);

                if (resBody.Contains(id)&&resBody.Contains(cname)&&resBody.Contains(usperson))
                {
                    test.Log(Status.Pass, "GetUSPersonLEIvalues Response body contains latest CounterpartyID: " + id + ", Name "+cname+ " & IsUSperson: "+usperson);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetUSPersonLEIvalues Response values doesn't contain " + id + " ,"+ cname + " &" + usperson);
                     Assert.Fail("USPersonLEIvalues Response is Failed!!");
                 }
            
            }
            else
            {
                test.Log(Status.Fail, "USPersonLEIvalues Response is " + resCode);
                Assert.Fail("USPersonLEIvalues Response code is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
