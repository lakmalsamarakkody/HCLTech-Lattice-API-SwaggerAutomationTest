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
    class ValidateUser
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='Login_Login_ValidateUser']/div[1]/h3/span[1]/a",
                         lattise_Try = "//*[@id='Login_Login_ValidateUser_content']/form/div[2]/input",
                         lattise_code = "//*[@id='Login_Login_ValidateUser_content']/div[2]/div[4]/pre",
                         inputD = "//div[@id='Login_Login_ValidateUser_content']/form/table/tbody/tr/td/textarea[@Name='userDetails']",
                        lattise_body = "//*[@id='Login_Login_ValidateUser_content']/div[2]/div[3]/pre/code";

        public ValidateUser(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement details => Driver.FindElement(By.XPath(inputD));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetValidateUser()
        {
            test = ExtReport.CreateTest("TS_079_GetValidateUser").Info("Test Started");

            string value = InputExcelAPI.GetCellData("WebAPI", 47, 3);

           /* string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 127, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 127, 4);*/


            SqlCommand cmd = new SqlCommand("select* from Trader where Id=77", con);
            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            
            reader.Read();

            string TraderID;
            string name;
            TraderID = Convert.ToString(reader[0]);
            name = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();
            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(123);
             var value = row.GetCell(3).NumericCellValue.ToString();
             var val2 = row.GetCell(4).StringCellValue;

             Thread.Sleep(2000);*/

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
            url_lattise.Click();
            test.Log(Status.Info, "GetValidateUser selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(inputD)));
            // details.SendKeys(@"['CORP\\l_samarakkody','DESKTOP-NJK2OUL','2.7.6.7']");
            string input = string.Format($"['{username}','{value}']");
            details.SendKeys(input);
            test.Log(Status.Info, "User Details entered");

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
                test.Log(Status.Pass, "GetValidateUser Response is " + resCode);

                if (resBody.Contains(TraderID)&&resBody.Contains(name))
                {
                    test.Log(Status.Pass, "GetValidateUser Response contains user ID: " +TraderID+" & Name: "+name);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetValidateUser Response is Failed!");
                    Assert.Fail("GetValidateUser Response is Failed!!");
                }
            }
            else
            {
                test.Log(Status.Fail, "GetValidateUser Response is " + resCode);
                Assert.Fail("GetValidateUser Response code is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);
        }
    }
}
