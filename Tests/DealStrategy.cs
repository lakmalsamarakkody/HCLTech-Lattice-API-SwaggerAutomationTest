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
    class DealStrategy
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();


        static string strat_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetDealStrategy']/div[1]/h3/span[1]/a",
                         strat_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetDealStrategy_content']/form/div[2]/input",
                         strat_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetDealStrategy_content']/div[2]/div[4]/pre",
                         uname = "//div[@id='DatabaseTopic_DatabaseTopic_GetDealStrategy_content']/form/table/tbody/tr/td/input[@name='Username']",
                         dealID = "//*[@id='DatabaseTopic_DatabaseTopic_GetDealStrategy_content']/form/table/tbody/tr/td/input[@name='dealId']",
                        strat_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetDealStrategy_content']/div[2]/div[3]/pre/code";

        public DealStrategy(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_strat => Driver.FindElement(By.XPath(strat_url));
        IWebElement button_strat => Driver.FindElement(By.XPath(strat_Try));
        IWebElement body_strat => Driver.FindElement(By.XPath(strat_body));
        IWebElement un => Driver.FindElement(By.XPath(uname));
        IWebElement dealId => Driver.FindElement(By.XPath(dealID));
        IWebElement res_strat => Driver.FindElement(By.XPath(strat_code));


        public void GetDealStrategy()
        {

            test = ExtReport.CreateTest("Get_DealStrategy").Info("DealStrategy Test has Started");

            string val = InputExcelAPI.GetCellData("WebAPI", 21, 3);
            //string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 30, 3);

            SqlCommand cmd = new SqlCommand("select * from Deal where Id = @Id", con);
            cmd.Parameters.AddWithValue("@Id", val);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string strategy;
            strategy = Convert.ToString(reader[10]);

            reader.Close();
            con.Close();

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(24);
             var val = row.GetCell(3).StringCellValue;
             //var val = row.GetCell(0).StringCellValue;
             //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
             Thread.Sleep(2000);*/

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(strat_url)));

            url_strat.Click();

            test.Log(Status.Info, "GCDID selected");
            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(dealID)));

            dealId.SendKeys(val);

            test.Log(Status.Info, "Deal ID entered");

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(uname)));

            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(strat_Try)));

            button_strat.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
            //  Thread.Sleep(3000);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(strat_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_strat);
            action.Perform();

            string resBody = body_strat.Text;
            string resCode = res_strat.Text;

            //test.Log(Status.Info, resBody);
            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "Deal Strategy Response is " + resCode);

                if (resBody.Contains(strategy))
                {
                    test.Log(Status.Pass, "Deal Strategy Response body contains Deal Strategy " +strategy);
                    test.Log(Status.Pass, "Test Passed!");
                }
                else
                {
                    test.Log(Status.Fail, "Deal Strategy Response is failed!! not contain " +strategy + ", Instead it contains " + resBody);
                    Assert.Fail("Deal Strategy Response is "+ resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "Deal Strategy Response is " + resCode);
                Assert.Fail("Deal Strategy Response is ", resCode);
            }

            url_strat.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);

        }
    }
}
