using AventStack.ExtentReports;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using SwaggerWebAPI.Libs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SwaggerWebAPI
{
    class Cutoff
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();

        static string cut_url = "//*[@id='Cutoffs_Cutoffs_GetCutoffs']/div[1]/h3/span[1]/a",
                         cut_Try = "//*[@id='Cutoffs_Cutoffs_GetCutoffs_content']/form/div[2]/input",
                         Res_code = "//*[@id='Cutoffs_Cutoffs_GetCutoffs_content']/div[2]/div[4]/pre",
                         Username = "//div[@id='Cutoffs_Cutoffs_GetCutoffs_content']/form/table/tbody/tr/td/input[@name='Username']",
                         cut_body = "//*[@id='Cutoffs_Cutoffs_GetCutoffs_content']/div[2]/div[3]/pre/code";


        public Cutoff(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_cut => Driver.FindElement(By.XPath(cut_url));
        IWebElement button_cut => Driver.FindElement(By.XPath(cut_Try));
        IWebElement body_cut => Driver.FindElement(By.XPath(cut_body));
        IWebElement un => Driver.FindElement(By.XPath(Username));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));


        public void GetCutoff()
        {

            test = ExtReport.CreateTest("TC_008_Cutoff").Info("Test Started");

            //  Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
           // WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(cut_url)));
            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(cut_url))).Click();

            //url_cut.Click();

            test.Log(Status.Info, "Cutoff selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Username)));

            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(cut_Try)));

            button_cut.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(cut_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_cut);
            action.Perform();

           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);

            string resBody = body_cut.Text;
            string resCode = res_Code.Text;
            
            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GetCutoffs Response is " + resCode);

                int getrowcount = InputExcelAPI.GetLastRowIndex("Cutoffs");
                int getcolumncount = InputExcelAPI.GetColumnCountByRow("Cutoffs", getrowcount);
                // IDictionary<string, string> numberNames = new Dictionary<string, string>();



                for (int rowval = 1; rowval <= getrowcount; rowval++)
                {
                    // get all details of row
                    List<string> dataset = new List<string>();
                    for (int colval = 0; colval < getcolumncount; colval++)
                    {
                        dataset.Add(InputExcelAPI.GetCellData("Cutoffs", rowval, colval));
                    }

                    // check all details are exist in response body
                    bool dataExists = true;
                    foreach (var data in dataset)
                    {
                        if (!resBody.Contains(data))
                        {
                            dataExists = false;
                        }
                    }



                    if (dataExists)
                    {
                        test.Log(Status.Pass, "Cutoff Response body contains ID " + dataset[0] + "Cutoff Time is" + dataset[1]);
                    }
                    else
                    {
                        test.Log(Status.Fail, "Cutoff Response is failed!! not contains " + dataset[0] + "& " + dataset[1]);
                        Assert.Fail("Response doesn't containg " + dataset[0]);
                    }
                }
                test.Log(Status.Pass, "Test 8 is Pass");



            }
            else
            {
                test.Log(Status.Fail, "Cutoff Response is " + resCode);
                Assert.Fail("Response code " + resCode);

            }

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(cut_url)));
            //url_cut.Click();
          //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);

        }
    }

}
