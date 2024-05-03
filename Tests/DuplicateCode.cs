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
    class DuplicateCode
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string url = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeDuplicateCodeForAdd']/div[1]/h3/span[1]/a",
                         Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeDuplicateCodeForAdd_content']/form/div[2]/input",
                         code = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeDuplicateCodeForAdd_content']/div[2]/div[4]/pre",
                        Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetLatticeDuplicateCodeForAdd_content']/form/table/tbody/tr/td/input[@name='Username']",
                        codeIn = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeDuplicateCodeForAdd_content']/form/table/tbody/tr/td/input[@name='code']",
                        body = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeDuplicateCodeForAdd_content']/div[2]/div[3]/pre/code";

        public DuplicateCode(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
        }

        IWebElement url_dup => Driver.FindElement(By.XPath(url));
        IWebElement button_dup => Driver.FindElement(By.XPath(Try));
        IWebElement body_dup => Driver.FindElement(By.XPath(body));
        IWebElement res_dup => Driver.FindElement(By.XPath(code));
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement inputC => Driver.FindElement(By.XPath(codeIn));
        //IWebElement inputC => Driver.FindElement(By.Name("code"));


        public void GetDuplicateCode()
        {

            test = ExtReport.CreateTest("TC_021_GetDuplicateCode").Info("DuplicateCode Test has Started");

            string inputCode = InputExcelAPI.GetCellData("WebAPI", 16, 3);
            /*string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 22, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 22, 4);
*/
            SqlCommand cmd = new SqlCommand("select * from Counterparty where Code = @Code", con);
            cmd.Parameters.AddWithValue("@Code", inputCode);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string name;
            string code;
            code = Convert.ToString(reader[1]);
            //name = Convert.ToString(reader[2]);

            reader.Close();
            con.Close();

            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(5);
            var row = sheet.GetRow(30);
            var val = row.GetCell(3).NumericCellValue.ToString();
            //var val = row.GetCell(0).StringCellValue;
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
            Thread.Sleep(2000);*/

            //WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(url)));
            // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(url)));

            url_dup.Click();

            test.Log(Status.Info, "GCDID selected");
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);

            //WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(codeIn)));
            Thread.Sleep(1000);

            inputC.SendKeys(inputCode);
            test.Log(Status.Info, "Code : "+inputCode+ " entered");

           // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            string input = string.Format(username);
            un.SendKeys(input);

            button_dup.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
            //Thread.Sleep(3000);

            // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(body)));
            Thread.Sleep(2000);

            Actions action = new Actions(Driver);
            action.MoveToElement(body_dup);
            action.Perform();

            string resBody = body_dup.Text;
            string resCode = res_dup.Text;

            //test.Log(Status.Info, resBody);
            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "DuplicateCode Response is " + resCode);

                if (resBody.Contains(code))
                {
                    test.Log(Status.Pass, "DuplicateCode Response body contains Code "+code);
                    test.Log(Status.Pass, "Test Passed!");
                }
                else
                {
                    test.Log(Status.Warning, "DuplicateCode Response is failed!! not contain " + code + ", Instead it contains " + resBody+ " No Duplicate Code!");
                    Assert.Fail("DuplicateCode Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "DuplicateCode Response is " + resCode);
                Assert.Fail("DuplicateCode Response code is " + resCode);
            }

            url_dup.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);

        }
    }
}
