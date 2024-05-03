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
using System.Threading.Tasks;

namespace SwaggerWebAPI
{
    class LastchangeIDofCounterparty
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;

        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();

        static string last_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetlastchangeIdOfCounterparty']/div[1]/h3/span[1]/a",
                         last_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetlastchangeIdOfCounterparty_content']/form/div[2]/input",
                         last_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetlastchangeIdOfCounterparty_content']/div[2]/div[4]/pre",
                         Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetlastchangeIdOfCounterparty_content']/form/table/tbody/tr/td/input[@name='Username']",
                        last_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetlastchangeIdOfCounterparty_content']/div[2]/div[3]/pre/code";

        public LastchangeIDofCounterparty(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_last => Driver.FindElement(By.XPath(last_url));
        IWebElement button_last => Driver.FindElement(By.XPath(last_Try));
        IWebElement body_last => Driver.FindElement(By.XPath(last_body));
        IWebElement un => Driver.FindElement(By.XPath(Username));

        IWebElement res_last => Driver.FindElement(By.XPath(last_code));


        public void GetLastchange()
        {

            test = ExtReport.CreateTest("TC_015_Lastchange").Info("Test Started");

            SqlCommand cmd = new SqlCommand("select * from CounterpartyChange where Id= (select max(Id) from CounterpartyChange)", con);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            string lastId;
            string code;

            reader.Read();

            lastId = Convert.ToString(reader[0]);
            code = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();

            /*
             string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

                        XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

                        var sheet = wb.GetSheetAt(5);
                        var row = sheet.GetRow(14);
                        var val = row.GetCell(3).NumericCellValue.ToString();*/
            //var val = row.GetCell(0).StringCellValue;
            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(last_url)));

            url_last.Click();

            test.Log(Status.Info, "Lastchange selected");
            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(last_Try)));

            button_last.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(last_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_last);
            action.Perform();

            string resBody = body_last.Text;
            string resCode = res_last.Text;

            //test.Log(Status.Info, resBody);
            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "Lastchange Response is " + resCode);

                if (resBody.Contains(lastId))
                {
                    test.Log(Status.Pass, "Lastchange Response body contains " +lastId+" & Code "+code);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "Lastchange Response is failed!! not contain " + lastId + ", Instead it contains " + resBody);
                    Assert.Fail("Lastchange Response contain " + resBody);

                }
            }
            else
            {
                test.Log(Status.Fail, "BrokerCodes Response is " + resCode);
                Assert.Fail("Lastchange Response contain " + resCode);

            }

            url_last.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }
    }
}
