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
    class LatestStatusChange
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='Status_Status_GetLatestStatusChange']/div[1]/h3/span[1]/a",
                         lattise_Try = "//*[@id='Status_Status_GetLatestStatusChange_content']/form/div[2]/input",
                         lattise_code = "//*[@id='Status_Status_GetLatestStatusChange_content']/div[2]/div[4]/pre",
                         userName = "//div[@id='Status_Status_GetLatestStatusChange_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='Status_Status_GetLatestStatusChange_content']/div[2]/div[3]/pre/code";


        public LatestStatusChange(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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


        public void GetLatestStatusChange()
        {
            test = ExtReport.CreateTest("TS_082_GetLatestStatusChange").Info("Test Started");

            SqlCommand cmd = new SqlCommand("select max(Id) from StatusChange", con);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            string lastId;
            //string deal;

            reader.Read();


            lastId = Convert.ToString(reader[0]);
            //deal = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(127);

             var value = row.GetCell(3).NumericCellValue.ToString();*/
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
            url_lattise.Click();
            test.Log(Status.Info, "GetLatestStatusChange selected");

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


            test.Log(Status.Info, "Verifying Value....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GetLatestStatusChange Response is " + resCode);

                if (resBody.Contains(lastId))
                {
                    test.Log(Status.Pass, "GetLatestStatusChange Response contains Latest Status change ID " + lastId);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetLatestStatusChange Response is Failed!");
                    Assert.Fail("GetLatestStatusChange Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "GetLatestStatusChange Response is " + resCode);
                Assert.Fail("GetLatestStatusChange Response code is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
