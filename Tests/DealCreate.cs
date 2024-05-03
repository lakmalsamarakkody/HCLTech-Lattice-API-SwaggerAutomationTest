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
    class DealCreate
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string deal_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetDealCreatedAndLastUpdatedDateAndTime']/div[1]/h3/span[1]/a",
                         deal_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetDealCreatedAndLastUpdatedDateAndTime_content']/form/div[2]/input",
                         Res_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetDealCreatedAndLastUpdatedDateAndTime_content']/div[2]/div[4]/pre",
                         Username = "//div[@id='DatabaseTopic_DatabaseTopic_GetDealCreatedAndLastUpdatedDateAndTime_content']/form/table/tbody/tr/td/input[@name='Username']",
                        deal_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetDealCreatedAndLastUpdatedDateAndTime_content']/div[2]/div[3]/pre/code";

        public DealCreate(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_deal => Driver.FindElement(By.XPath(deal_url));
        IWebElement button_deal => Driver.FindElement(By.XPath(deal_Try));
        IWebElement body_deal => Driver.FindElement(By.XPath(deal_body));
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement dealField => Driver.FindElement(By.Name("dealid"));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
      //  IWebElement dealField => Driver.FindElement(By.XPath(dealIDfield));

        public void GetDealCreate()
        {

            test = ExtReport.CreateTest("TS_013_GetDealCreatedandLastUpdate").Info("Test Started");

            string val1 = InputExcelAPI.GetCellData("WebAPI", 7, 3);
            string val = ValidationExcelAPI.GetCellData("DatabaseTopic", 10, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 10, 4);
            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(10);
             var val = row.GetCell(3).StringCellValue;
             var val2 = row.GetCell(4).StringCellValue;
             Thread.Sleep(2000);*/
            //WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(deal_url)));

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(deal_url)));

            url_deal.Click();
        
            test.Log(Status.Info, "DealCreate selected");
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

            dealField.SendKeys(val1);
            test.Log(Status.Info, "DealID : 34 has entered");

            string input = string.Format(username);
            un.SendKeys(input);

            button_deal.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(deal_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_deal);
            action.Perform();

            string resBody = body_deal.Text;
            string resCode = res_Code.Text;

            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "DealCreate Response is " + resCode);

                if (resBody.Contains(val)&& resBody.Contains(val2))
                {
                    test.Log(Status.Pass, "DealCreate Response body contains Create Date " +val+" & Update Date "+val2);
                    test.Log(Status.Pass, "DealCreate Test  is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "DealCreate Response is failed!! not caontain "+val+" &"+val2+ ", Instead it contains " + resBody);
                    Assert.Fail("DealCreate Response contain "+resBody );

                }
            }
            else
            {
                test.Log(Status.Fail, "DealCreate Response is " + resCode);
                Assert.Fail("DealCreate Response is " + resCode);

            }

            url_deal.Click();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
