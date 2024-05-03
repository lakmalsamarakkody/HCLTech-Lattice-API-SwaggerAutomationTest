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
    class USPerson
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetUSPersonValue']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetUSPersonValue_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetUSPersonValue_content']/div[2]/div[4]/pre",
                         counter = "//div[@id='DatabaseTopic_DatabaseTopic_GetUSPersonValue_content']/form/table/tbody/tr/td/input[@name='counterpartyID']",
                        userName = "//*[@id='DatabaseTopic_DatabaseTopic_GetUSPersonValue_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetUSPersonValue_content']/div[2]/div[3]/pre/code";


        public USPerson(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement counterparty => Driver.FindElement(By.XPath(counter));
        IWebElement un => Driver.FindElement(By.XPath(userName));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetUSPersonValue()
        {

            test = ExtReport.CreateTest("Get_GetUSPersonValue").Info("Test Started");

            string value = InputExcelAPI.GetCellData("WebAPI", 34, 3);

            // string val = ValidationExcelAPI.GetCellData("DatabaseTopic", 60, 3);

            //SqlCommand cmd = new SqlCommand("select * from Counterparty where ID = (select max(ID) from Counterparty)", con);
            SqlCommand cmd = new SqlCommand("select * from Counterparty where Code = @Code", con);

            cmd.Parameters.AddWithValue("@Code", value);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string code;
            string usperson,id;
           
            code = Convert.ToString(reader[1]).Trim();
            id = Convert.ToString(reader[0]);
            usperson = Convert.ToString(reader[8]).Trim();

            reader.Close();
            con.Close();

            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(5);
            var row = sheet.GetRow(46);
            var value = row.GetCell(3).StringCellValue;
            Thread.Sleep(2000);*/

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
            url_lattise.Click();
            test.Log(Status.Info, "GetUSPersonValue selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(counter)));
            counterparty.SendKeys(id);
            test.Log(Status.Info, "CounterpartyID : "+id+ " has entered");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_Try)));
            button_lattise.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_lattise);
            action.Perform();

            string resBody = body_lattise.Text;
            string resCode = res_Code.Text;


            test.Log(Status.Info, "Verifying Value....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GetUSPersonValue Response is " + resCode);

                if (resBody.Contains(usperson))
                {
                    test.Log(Status.Pass, "GetUSPersonValue Response body contains Counterparty Code: "+code+ " & person "+usperson);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "USPersonValue Response is fail");
                    Assert.Fail("USPersonValue Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "USPersonValue Response is " + resCode);
                Assert.Fail("USPersonValue Response is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
