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
    class FixingSource
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string fix_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetFixingSource']/div[1]/h3/span[1]/a",
                         fix_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetFixingSource_content']/form/div[2]/input",
                         Res_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetFixingSource_content']/div[2]/div[4]/pre",
                        Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetFixingSource_content']/form/table/tbody/tr/td/input[@name='Username']",
                        fix_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetFixingSource_content']/div[2]/div[3]/pre/code";

        public FixingSource(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_fix => Driver.FindElement(By.XPath(fix_url));
        IWebElement button_fix => Driver.FindElement(By.XPath(fix_Try));
        IWebElement body_fix => Driver.FindElement(By.XPath(fix_body));
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));


        public void GetFixSource()
        {

            test = ExtReport.CreateTest("TS_011_GetFixingSource").Info("FixingSource Test is Started");

            string val = ValidationExcelAPI.GetCellData("DatabaseTopic", 8, 3);

            SqlCommand cmd = new SqlCommand("select * from FixingSource where Id = @Id", con);
            cmd.Parameters.AddWithValue("@Id", val);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            string name;

            reader.Read();

            name = Convert.ToString(reader[1]);


            // code = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();



            //WorkBook wb = WorkBook.Load(@"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx");

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(1);
             var val = row.GetCell(4).StringCellValue;*/

            // var val2 = row.GetCell(0).StringCellValue;
            // var val = row.GetCell(0).StringCellValue.Trim();
            // Console.WriteLine(val);
            // WorkSheet ws = wb.GetWorkSheet("sheet1");


            //Console.WriteLine(val);
            //  WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(fix_url)));

            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(fix_url)));

            url_fix.Click();

            test.Log(Status.Info, "Fixing source selected");
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            string input = string.Format(username);
            un.SendKeys(input);

            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);
           
            button_fix.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(fix_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_fix);
            action.Perform();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);

            string resBody = body_fix.Text;
            string resCode = res_Code.Text;

            // Console.WriteLine(resBody);

            // test.Log(Status.Info, resBody);
            test.Log(Status.Info, "Verifying Fixing source Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "Fixing source Response is " + resCode);
                 
                if (resBody.Contains(name))
                {
                    test.Log(Status.Pass, "Fixing source Response body contains " + name);
                    test.Log(Status.Pass, "Fixingsource verification is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "Conversion Response body verification is failed!! " + val + " is not containing inside the response body");
                    Assert.Fail("Fixingsource Response is failed!");
                }
            }
        

            else
            {
                test.Log(Status.Fail, "Fixingsource Response is " + resCode);
                Assert.Fail("Fixingsource Response is ", resCode);

            }

            url_fix.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }
    }
}
