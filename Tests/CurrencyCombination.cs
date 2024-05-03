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
    class CurrencyCombination
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();


        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_CheckCurrencyCombinationExists']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_CheckCurrencyCombinationExists_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_CheckCurrencyCombinationExists_content']/div[2]/div[4]/pre",
                         inputD = "//div[@id='DatabaseTopic_DatabaseTopic_CheckCurrencyCombinationExists_content']/form/table/tbody/tr/td/input[@name='existingCurrencyCombination']",
                        userName = "//*[@id='DatabaseTopic_DatabaseTopic_CheckCurrencyCombinationExists_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_CheckCurrencyCombinationExists_content']/div[2]/div[3]/pre/code";


        public CurrencyCombination(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement combo => Driver.FindElement(By.XPath(inputD));
        IWebElement un => Driver.FindElement(By.XPath(userName));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetCurrencyCombination()
        {
            test = ExtReport.CreateTest("TS_058_GetCurrencyCombination").Info("Test Started");

            string val = InputExcelAPI.GetCellData("WebAPI", 39, 3);

            SqlCommand cmd = new SqlCommand("select * from CcyPair where Code = @Code", con);

            cmd.Parameters.AddWithValue("@Code", val);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string code;

            code = Convert.ToString(reader[1]).Trim();

            reader.Close();
            con.Close();

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "GetCurrencyCombination selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(inputD)));

            combo.SendKeys(val);
            test.Log(Status.Info, "Currency combo : " + val+ " has entered");

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
                test.Log(Status.Pass, "GetCurrencyCombination Response is " + resCode);

                if (resBody == "true" && code==val)
                {
                    test.Log(Status.Pass, "GetCurrencyCombination Response is True, " +val+ " currency pair exists");
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetCurrencyCombination Response is False!, currency pair doesen't exist");
                    Assert.Fail("GetCurrencyCombination Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "GetCurrencyCombination Response  code is " + resCode);
                Assert.Fail("GetCurrencyCombination Response code is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
