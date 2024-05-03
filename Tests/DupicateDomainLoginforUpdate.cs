using AventStack.ExtentReports;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SwaggerWebAPI.Libs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace SwaggerWebAPI
{
    class DupicateDomainLoginforUpdate
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeDuplicateDomainLoginForUpdate']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeDuplicateDomainLoginForUpdate_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeDuplicateDomainLoginForUpdate_content']/div[2]/div[4]/pre",
                         data = "//div[@id='DatabaseTopic_DatabaseTopic_GetLatticeDuplicateDomainLoginForUpdate_content']/form/table/tbody/tr/td/textarea[@name='data']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetLatticeDuplicateDomainLoginForUpdate_content']/div[2]/div[3]/pre/code";


        public DupicateDomainLoginforUpdate(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement  inputD => Driver.FindElement(By.XPath(data));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void Get_DupicateDomainLoginforUpdate()
        {
            test = ExtReport.CreateTest("TS_049_GetDupicateDomainLoginforUpdate").Info("Test Started");

            string value = InputExcelAPI.GetCellData("WebAPI", 37, 3);

            string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 66, 3);

            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(5);
            var row = sheet.GetRow(52);
            var value = row.GetCell(3).StringCellValue;

            Thread.Sleep(2000);*/

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "DupicateDomainLoginforUpdate selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(data)));

            //inputD.SendKeys(@"{'70':'CORP\\l_samarakkody'}");
            //'70':'CORP\\l_samarakkody','-1':
            // { '70':'CORP\\l_samarakkody','-1':'Q09SUFxqX2dhbWFnZQ=='}
            string input = "{" + value + "'" + username + "'}";
            //string  = string.Format(@"{1}'{0}'", username, value);
             //string input = string.Format($"['{value}','{username}']");
            inputD.SendKeys(input);
            test.Log(Status.Info, "Data : "+value+ " entered");

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
                test.Log(Status.Pass, "DupicateDomainLoginforUpdate Response is " + resCode);

                if (resBody.Contains(val1))
                {
                    test.Log(Status.Pass, "DupicateDomainLoginforUpdate Response contains " + val1);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else if (resBody == "{}")
                {
                    test.Log(Status.Pass, "DupicateDomainLoginforUpdate Response doesn't have any duplicate domain logins");
                    test.Log(Status.Pass, "Test is Pass");
                }

                else
                {
                    test.Log(Status.Fail, "DupicateDomainLoginforUpdate Response is Failed!");
                    Assert.Fail("DupicateDomainLoginforUpdate Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "DupicateDomainLoginforUpdate Response is " + resCode);
                Assert.Fail("DupicateDomainLoginforUpdate Response is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
