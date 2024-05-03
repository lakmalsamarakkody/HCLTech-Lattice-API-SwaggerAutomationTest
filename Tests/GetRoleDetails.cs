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
    class GetRoleDetails
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetRoleDetails']/div[1]/h3/span[1]/a",
        //static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetRoleDetails']/div[1]/h3/span[1]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetRoleDetails_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetRoleDetails_content']/div[2]/div[4]/pre",
                         lst = "//div[@id='DatabaseTopic_DatabaseTopic_GetRoleDetails_content']/form/table/tbody/tr/td/textarea[@name='lstresult']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetRoleDetails_content']/div[2]/div[3]/pre/code";
        //*[@id="DatabaseTopic_DatabaseTopic_GetRoleDetails_content"]/div[2]/div[3]/pre/code

        public GetRoleDetails(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
        }

        IWebElement url_lattise => Driver.FindElement(By.XPath(lattise_url));
        IWebElement button_lattise => Driver.FindElement(By.XPath(lattise_Try));
        IWebElement body_lattise => Driver.FindElement(By.XPath(lattise_body));
        IWebElement input => Driver.FindElement(By.XPath(lst));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));

        public void RoleDetails()
        {
            test = ExtReport.CreateTest("TS_021_GetRoleDetails").Info("Test Started");
            string val1 = InputExcelAPI.GetCellData("WebAPI", 15, 3);

            //WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath(lattise_url)));
           // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lattise_url)));
           // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
            url_lattise.Click();
            test.Log(Status.Info, "GetRoleDetails selected");

            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(4000);

            string input1 = string.Format(@"['{1}','{0}']", username, val1);
            input.SendKeys(input1);

          //  WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_Try)));
            button_lattise.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            //            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lattise_body)));
            Thread.Sleep(2000);
            Actions action = new Actions(Driver);
            action.MoveToElement(body_lattise);
            action.Perform();

            string resBody = body_lattise.Text;
            string resCode = res_Code.Text;


            test.Log(Status.Info, "Verifying Value....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GetRoleDetails Response is " + resCode);

                if (resBody == "true")
                {
                    test.Log(Status.Pass, "GetRoleDetails Response is True");
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetRoleDetails Response is False");
                    Assert.Fail("GetRoleDetails Response is ", resCode);
                }
            }
            else
            {
                test.Log(Status.Fail, "GetRoleDetails Response is " + resCode);
                Assert.Fail("GetRoleDetails Response is ", resCode);

            }

            url_lattise.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
