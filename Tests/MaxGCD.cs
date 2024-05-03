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
    class MaxGCD
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;

        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();

        static string gcd_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetMaxGCDId']/div[1]/h3/span[1]/a",
                         gcd_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetMaxGCDId_content']/form/div[2]/input",
                         gcd_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetMaxGCDId_content']/div[2]/div[4]/pre",
                         Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetMaxGCDId_content']/form/table/tbody/tr/td/input[@name='Username']",
                        gcd_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetMaxGCDId_content']/div[2]/div[3]/pre/code";

        public MaxGCD(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_gcd => Driver.FindElement(By.XPath(gcd_url));
        IWebElement button_gcd => Driver.FindElement(By.XPath(gcd_Try));
        IWebElement body_gcd => Driver.FindElement(By.XPath(gcd_body));
        IWebElement un => Driver.FindElement(By.XPath(Username));

        IWebElement res_gcd => Driver.FindElement(By.XPath(gcd_code));


        public void GetMaxGCD()
        {

            test = ExtReport.CreateTest("TC_016_MaxGCD").Info("MaxGCD Test has Started");

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(1);
             var val = row.GetCell(18).NumericCellValue.ToString();*/
            //var val = row.GetCell(0).StringCellValue;
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
          //WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(gcd_url)));
          WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(gcd_url)));

            url_gcd.Click();

            test.Log(Status.Info, "MaxGCD selected");
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(gcd_Try)));

            button_gcd.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");


            SqlCommand cmd = new SqlCommand("select max(GcdId) from Counterparty", con);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            int maxId;

            reader.Read();

            maxId = (int)reader[0];
            maxId += 1;
          
           // code = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(gcd_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_gcd);
            action.Perform();

            string resBody = body_gcd.Text;
            string resCode = res_gcd.Text;

            //test.Log(Status.Info, resBody);
            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "MaxGCD Response is " + resCode);

                if (resBody.Contains(maxId.ToString()))
                {
                    test.Log(Status.Pass, "MaxGCD Response body contains " + maxId);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "MaxGCD Response is failed!! not contain " + maxId + ", Instead it contains "+resBody);
                    Assert.Fail("MaxGCD Response contain " +resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "MaxGCD Response is " + resCode); 
                Assert.Fail("MaxGCD Response contain " + resCode);
            }

            url_gcd.Click();

        }
    }
}
