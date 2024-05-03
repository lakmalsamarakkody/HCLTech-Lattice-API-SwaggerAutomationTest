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
    class DealChanges
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='Deal_Deal_GetDealChanges']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='Deal_Deal_GetDealChanges_content']/form/div[2]/input",
                         lattise_code = "//*[@id='Deal_Deal_GetDealChanges_content']/div[2]/div[4]/pre",
                         inputD = "//div[@id='Deal_Deal_GetDealChanges_content']/form/table/tbody/tr/td/input[@name='_lastChangeId']",
                         userName = "//div[@id='Deal_Deal_GetDealChanges_content']/form/table/tbody/tr/td/input[@name='Username']",
                         lattise_body = "//*[@id='Deal_Deal_GetDealChanges_content']/div[2]/div[3]/pre/code";


        public DealChanges(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement id => Driver.FindElement(By.XPath(inputD));
        IWebElement UN => Driver.FindElement(By.XPath(userName));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetDealChanges()
        {
            test = ExtReport.CreateTest("TS_078_GetDealChanges").Info("Test Started");

           // string value = InputExcelAPI.GetCellData("WebAPI", 20, 3);

           /* string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 28, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 28, 4);
            string val3 = ValidationExcelAPI.GetCellData("DatabaseTopic", 28, 5);*/


            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(105);
             var value = row.GetCell(3).NumericCellValue.ToString();
             var val2 = row.GetCell(4).NumericCellValue.ToString();
             var val3 = row.GetCell(5).NumericCellValue.ToString();
             var val4 = row.GetCell(6).NumericCellValue.ToString();
             Thread.Sleep(2000);*/

            SqlCommand cmd = new SqlCommand("select max(Id) from DealChange where Id < (select max(Id) from DealChange)", con);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            reader.Read();

            string lastID;
            lastID = Convert.ToString(reader[0]);
            reader.Close();

            SqlCommand cmd1 = new SqlCommand("select Id,DealID from DealChange where Id = (select max(Id) from DealChange)", con);
            SqlDataReader reader2 = cmd1.ExecuteReader();
            reader2.Read();

            string maxID;
            string dealId;
            maxID = Convert.ToString(reader2[0]);
            dealId = Convert.ToString(reader2[1]);

            reader2.Close();
            con.Close();

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
            url_lattise.Click();
            test.Log(Status.Info, "GetDealChanges selected");


            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(inputD)));
            id.SendKeys(lastID);
            test.Log(Status.Info, "Latest Changed Deal ID: "+lastID+ "entered");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));
            string input = string.Format(username);
            UN.SendKeys(username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lattise_Try)));
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
                test.Log(Status.Pass, "GetDealChanges Response is " + resCode);

                if (resBody.Contains(maxID)&&resBody.Contains(dealId))
                {
                    test.Log(Status.Pass, "GetDealChanges Response contains DealID " + maxID + "& LastChangedDealID " + dealId);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetDealChanges Response is Failed!");
                    Assert.Fail("GetDealChanges Response is Failed!!");
                }
            }
            else
            {
                test.Log(Status.Fail, "GetDealChanges Response is " + resCode);
                Assert.Fail("GetDealChanges Response code is " + resCode);
            }

            url_lattise.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
