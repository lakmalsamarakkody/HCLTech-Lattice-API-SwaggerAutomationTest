using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using System.Threading;
using AventStack.ExtentReports;
using NPOI.XSSF.UserModel;
using System.IO;
using OpenQA.Selenium.Interactions;
using SwaggerWebAPI.Libs;
using OpenQA.Selenium.Support.UI;
using System.Data.SqlClient;
using NUnit.Framework;

namespace SwaggerWebAPI
{
    class Counterparty
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string Counter_url = "//*[@id='Counterparty_Counterparty_GetCounterparties']/div[1]/h3/span[1]/a",
                         Counter_Try = "//*[@id='Counterparty_Counterparty_GetCounterparties_content']/form/div[2]/input",
                         Res_code = "//*[@id='Counterparty_Counterparty_GetCounterparties_content']/div[2]/div[4]/pre",
                        userName = "//*[@id='Counterparty_Counterparty_GetCounterparties_content']/form/table/tbody/tr/td/input[@name='Username']",
                        Counter_body = "//*[@id='Counterparty_Counterparty_GetCounterparties_content']/div[2]/div[3]/pre/code";

        public Counterparty(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }
    
        IWebElement url_Counter => Driver.FindElement(By.XPath(Counter_url));
        IWebElement button_Counter => Driver.FindElement(By.XPath(Counter_Try));
        IWebElement body_Counter => Driver.FindElement(By.XPath(Counter_body));
        IWebElement un => Driver.FindElement(By.XPath(userName));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));


        public void GetCounterparty()
        {

            test = ExtReport.CreateTest("TS_004_GetCounterparty").Info("Test Started");

           // string val1 = ValidationExcelAPI.GetCellData("CounterParty", 1, 0);
           // string val2 = ValidationExcelAPI.GetCellData("CounterParty", 1, 1);

            SqlCommand cmd = new SqlCommand("select * from Counterparty where Id = (select max(Id) from Counterparty)", con);
         
            //cmd.Parameters.AddWithValue("@Code", val1);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string name;
            string code;
            code = Convert.ToString(reader[1]);
            name = Convert.ToString(reader[2]);

            reader.Close();
            con.Close();

            /* WorkBook wb = WorkBook.Load(@"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx");

             WorkSheet ws = wb.GetWorkSheet("sheet1");

             String value = (string)ws["A3"].Value;*/

            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(2);
            var row = sheet.GetRow(1);
            var val = row.GetCell(0).StringCellValue;

            Thread.Sleep(2000);*/

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(Counter_url)));
            url_Counter.Click();
            test.Log(Status.Info, "Counterparty selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(Counter_Try)));
            button_Counter.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Counter_body)));
            Actions action = new Actions(Driver);
            action.MoveToElement(body_Counter);
            action.Perform();

            string resBody = body_Counter.Text;
            string resCode = res_Code.Text;

            //test.Log(Status.Info, resBody);
            test.Log(Status.Info, "Verifying Counterparty Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "Counterparty Response is " + resCode);

                if (resBody.Contains(code)&& resBody.Contains(name))
                {
                    test.Log(Status.Pass, "Counterparty Response body contains Code " +code+ " & Name "+name);
                    test.Log(Status.Pass, "Test 3 is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "Counterparty Response is fail");
                    Assert.Fail("Counterparty Response is Failed!!");
                }
            }
            else
            {
                test.Log(Status.Fail, "Counterparty Response is " + resCode);
                Assert.Fail("Counterparty Response code is " + resCode);
            }

            url_Counter.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);
        }
    }
}
