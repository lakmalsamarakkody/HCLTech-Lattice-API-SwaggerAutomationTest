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
    class GCDIDForEnteringOrUpdatingUserContactRef
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI,ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetGCDIDForEnteringOrUpdatingUserContactRef']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetGCDIDForEnteringOrUpdatingUserContactRef_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetGCDIDForEnteringOrUpdatingUserContactRef_content']/div[2]/div[4]/pre",
                         userName = "//*[@id='DatabaseTopic_DatabaseTopic_GetGCDIDForEnteringOrUpdatingUserContactRef_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetGCDIDForEnteringOrUpdatingUserContactRef_content']/div[2]/div[3]/pre/code";

        public GCDIDForEnteringOrUpdatingUserContactRef(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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

        public void GetUpdatingUserContactRef()
        {
            test = ExtReport.CreateTest("TS_043_GetGCDIDForEnteringOrUpdatingUserContactRef").Info("Test Started");

            //string value = ValidationExcelAPI.GetCellData("DatabaseTopic", 56, 3);

            SqlCommand cmd = new SqlCommand("select * from Trader as t where t.DomainLogin = @username", con);

            cmd.Parameters.AddWithValue("@username", User.GetUsername());

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string domain;
            string gcd;

            domain = Convert.ToString(reader[2]).Trim();
            gcd = Convert.ToString(reader[4]);

            reader.Close();
            con.Close();

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(52);
             var value = row.GetCell(3).NumericCellValue.ToString();
             //var val2 = row.GetCell(4).StringCellValue;
             Thread.Sleep(2000);*/

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "GetUpdatingUserContactRef selected");

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
                test.Log(Status.Pass, "GetUpdatingUserContactRef Response is " + resCode);

                if (resBody.Contains(gcd))
                {
                    test.Log(Status.Pass, "GetUpdatingUserContactRef Response body contains Domain: " +domain+ " & GCDID: "+gcd);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetUpdatingUserContactRef Response is Failed");
                    Assert.Fail("GetUpdatingUserContactRef Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "GetUpdatingUserContactRef Response is " + resCode);
                Assert.Fail("GetUpdatingUserContactRef Response is " + resCode);
            }

            url_lattise.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
