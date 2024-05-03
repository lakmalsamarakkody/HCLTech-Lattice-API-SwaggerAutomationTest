using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SwaggerWebAPI.Libs;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SwaggerWebAPI
{
    class GetAllDirections
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetAllDirections']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetAllDirections_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetAllDirections_content']/div[2]/div[4]/pre",
                         userName = "//*[@id='DatabaseTopic_DatabaseTopic_GetAllDirections_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetAllDirections_content']/div[2]/div[3]/pre/code";


        public GetAllDirections(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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


        public void Get_AllDirections()
        {

            test = ExtReport.CreateTest("TS_053_GetAllDirections").Info("Test Started");

            //string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 74, 3);

            SqlCommand cmd = new SqlCommand("select Direction from TradeChange", con);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string code;

            code = Convert.ToString(reader[0]).Trim();

            reader.Close();
            con.Close();
            // string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 68, 4);

            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(5);
            var row = sheet.GetRow(54);
            var value = row.GetCell(3).NumericCellValue.ToString();
            var val2 = row.GetCell(4).StringCellValue;
            Thread.Sleep(2000);*/

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "GetAllCcypairs selected");

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
                test.Log(Status.Pass, "GetAllCcypairs Response is " + resCode);

                if (resBody.Contains(code))
                {
                    test.Log(Status.Pass, "GetAllCcypairs Response body contains Direction " + code);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetAllCcypairs Response is Failed");
                    Assert.Fail("GetAllCcypairs Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "GetAllCcypairs Response is " + resCode);
                Assert.Fail("GetAllCcypairs Response is " + resCode);
            }

            url_lattise.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }

}
