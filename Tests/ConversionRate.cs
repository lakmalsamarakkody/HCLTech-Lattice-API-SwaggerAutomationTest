using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using SwaggerWebAPI.Libs;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Threading;

namespace SwaggerWebAPI.Tests
{
    class ConversionRate
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest ExtTest;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string Conversion_url ="//*[@id='ConversionRate_ConversionRate_GetUSDConversionRates']/div[1]/h3/span[2]/a",
                        Con_Try = "//*[@id='ConversionRate_ConversionRate_GetUSDConversionRates_content']/form/div[2]/input",
                        Res_code = "//*[@id='ConversionRate_ConversionRate_GetUSDConversionRates_content']/div[2]/div[4]/pre",
                       Con_allbody = "//*[@id='ConversionRate_ConversionRate_GetUSDConversionRates_content']/div[2]/div[3]/pre/code",
                       Conv_rate = " //*[@id='ConversionRate_ConversionRate_GetUSDConversionRates_content']/div[2]/div[3]/pre/code/span[2]",
                       C_username = "//*[@id='ConversionRate_ConversionRate_GetUSDConversionRates_content']/form/table/tbody/tr/td/input[@name='Username']";

        public ConversionRate(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement url_Conv => Driver.FindElement(By.XPath(Conversion_url));
        IWebElement button_Con => Driver.FindElement(By.XPath(Con_Try));
        // IWebElement body_Con => Driver.FindElement(By.XPath(Con_body));
        IWebElement allbody_Con => Driver.FindElement(By.XPath(Con_allbody));
        IWebElement res_Code => Driver.FindElement(By.XPath(Res_code));
        IWebElement res_rate => Driver.FindElement(By.XPath(Conv_rate));
        IWebElement UN => Driver.FindElement(By.XPath(C_username));

        public void GetConversionRate()
        {

            SqlCommand cmd = new SqlCommand("select * from ConversionRate where Id = (select max(Id) from ConversionRate)", con);
           
            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            string TradeID;
            string ConversionRate;

            reader.Read();

            TradeID = Convert.ToString(reader[2]);
            ConversionRate = Convert.ToString(reader[1]);

            reader.Close();

            con.Close();

            ExtTest = ExtReport.CreateTest("TS_003_GetConversionRate").Info("Test Started");
           // string value = ValidationExcelAPI.GetCellData("ConversionRate", 1, 0);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(Conversion_url))).Click();
//           url_Conv.Click();
            ExtTest.Log(Status.Info, "ConversionRate API Call selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(C_username)));
            UN.SendKeys(username);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(Con_Try)));
            button_Con.Click();

            ExtTest.Log(Status.Info, "Try it Now Button Clicked");
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(Con_allbody)));

            Actions action = new Actions(Driver);
            action.MoveToElement(allbody_Con);
            action.Perform();

            string resBody = allbody_Con.Text;
            string resCode = res_Code.Text;

            ExtTest.Log(Status.Info, "Verifying ConversionRate Values....");

            if (resCode == "200")
            {
                ExtTest.Log(Status.Pass, "ConversionRate Response is " + resCode);

                if (resBody.Contains(TradeID))
                {
                    ExtTest.Log(Status.Pass, "ConversionRate Response body contains Trade Id " + TradeID +" conversionrate : "+ ConversionRate);
                    ExtTest.Log(Status.Pass, "Test case is Pass");
                }
                else
                {
                    ExtTest.Log(Status.Fail, "ConversionRate Response is fail");
                    Assert.Fail("ConversionRate Response Failed ! ");
                }
            }
            else
            {
                ExtTest.Log(Status.Fail, "ConversionRate Response is " + resCode);
                Assert.Fail("ConversionRate Response is "+ resCode);
            }

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(Conversion_url))).Click();

           // url_Conv.Click();
           Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }




    }
}
