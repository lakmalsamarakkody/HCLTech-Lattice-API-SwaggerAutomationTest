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
    class GetGCDid
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();

        static string gcd_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetGCDId']/div[1]/h3/span[1]/a",
                         gcd_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetGCDId_content']/form/div[2]/input",
                         gcd_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetGCDId_content']/div[2]/div[4]/pre",
                        Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetGCDId_content']/form/table/tbody/tr/td/input[@name='Username']",
                       gcdId = "//div[@id='DatabaseTopic_DatabaseTopic_GetGCDId_content']/form/table/tbody/tr/td/input[@name='gcdId']",
                        gcd_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetGCDId_content']/div[2]/div[3]/pre/code";

        public GetGCDid(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
        }

        IWebElement url_gcd => Driver.FindElement(By.XPath(gcd_url));
        IWebElement button_gcd => Driver.FindElement(By.XPath(gcd_Try));
        IWebElement body_gcd => Driver.FindElement(By.XPath(gcd_body));
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement gcdld => Driver.FindElement(By.XPath(gcdId));
        IWebElement res_gcd => Driver.FindElement(By.XPath(gcd_code));


        public void GetGCDID()
        {

            test = ExtReport.CreateTest("TC_022_GetGCDId").Info("GCDID Test has Started");

            string value = ValidationExcelAPI.GetCellData("DatabaseTopic", 24, 3);
            string val1 = InputExcelAPI.GetCellData("WebAPI", 17, 3);

            SqlCommand cmd = new SqlCommand("select * from Counterparty where Code = @Code", con);
            cmd.Parameters.AddWithValue("@Code", val1);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string name;
            string code;
            code = Convert.ToString(reader[1]);
            name = Convert.ToString(reader[2]);

            reader.Close();
            con.Close();

            // string val = ExcelAPI.GetCellData("DatabaseTopic", 17, 3);
            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(22);
             var val = row.GetCell(3).StringCellValue;
             //var val = row.GetCell(0).StringCellValue;
             //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
             Thread.Sleep(2000);*/
            //WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(gcd_url)));

            url_gcd.Click();

            test.Log(Status.Info, "GCDID selected");
            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);

            //            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(gcdId)));
//            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(gcdId)));
            gcdld.SendKeys(val1);
            test.Log(Status.Info, "GCDID "+val1+" entered");

          //  WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));
            string input = string.Format(username);
            un.SendKeys(input);

          //  WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(gcd_Try)));
            button_gcd.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
            //Thread.Sleep(3000);

            //   WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(gcd_body)));

            Thread.Sleep(2000);

            Actions action = new Actions(Driver);
            action.MoveToElement(body_gcd);
            action.Perform();

            string resBody = body_gcd.Text;
            string resCode = res_gcd.Text;

            //test.Log(Status.Info, resBody);
            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GCDID Response is " + resCode);

                if (resBody.Contains(value))
                {
                    test.Log(Status.Pass, "GCDID Response body contains " + value);
                    test.Log(Status.Pass, "Test Passed!");
                }
                else
                {
                    test.Log(Status.Fail, "GCDID Response is failed!! not contain " + value);
                    Assert.Fail("GCDID Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "GCDID Response is " + resCode);
                Assert.Fail("GCDID Response code is " + resCode);
            }

            url_gcd.Click();
          //  Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

        }
    }
}
