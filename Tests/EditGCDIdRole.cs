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
    class EditGCDIdRole
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetEditGcdIdRoleForUser']/div[1]/h3/span[1]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetEditGcdIdRoleForUser_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetEditGcdIdRoleForUser_content']/div[2]/div[4]/pre",
                         lst = "//div[@id='DatabaseTopic_DatabaseTopic_GetEditGcdIdRoleForUser_content']/form/table/tbody/tr/td/textarea[@name='lstresult']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetEditGcdIdRoleForUser_content']/div[2]/div[3]/pre/code";

        public EditGCDIdRole(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement lstresult => Driver.FindElement(By.XPath(lst));

        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetEditGCDIdRole()
        {
            test = ExtReport.CreateTest("Get_EditGCDIdRole").Info("Test Started");

            string val1 = InputExcelAPI.GetCellData("WebAPI", 14, 3);

            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(1000);
          //WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lattise_url)));
          WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));
          url_lattise.Click();
            test.Log(Status.Info, "GetEditGCDIdRole selected");
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(1000);

            //string input = string.Format("['CORP','l_samarakkody','{0}']", username);
            string input = string.Format(@"['{1}','{0}']", username, val1);

            //WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lst)));

            lstresult.SendKeys(input);

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
                test.Log(Status.Pass, "GetEditGCDIdRole Response is " + resCode);

                if (resBody == "true")
                {
                    test.Log(Status.Pass, "GetEditGCDIdRole Response is True");
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Warning, "GetEditGCDIdRole Response is False");
                    Assert.Fail("Fixingsource Response is " +resBody);

                }
            }
            else
            {
                test.Log(Status.Fail, "GetEditGCDIdRole Response is " + resCode);
                Assert.Fail("Fixingsource Response is ", resCode);

            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

        }
    }
}
