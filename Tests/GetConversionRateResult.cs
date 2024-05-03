using AventStack.ExtentReports;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
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
    class GetConversionRateResult
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetConversionRateResult']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetConversionRateResult_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetConversionRateResult_content']/div[2]/div[4]/pre",
                         uname = "//*[@id='DatabaseTopic_DatabaseTopic_GetConversionRateResult_content']/form/table/tbody/tr/td/input[@name='Username']",
                        result =  "//*[@id='DatabaseTopic_DatabaseTopic_GetConversionRateResult_content']/form/table/tbody/tr/td/input[@name='tradeid']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetConversionRateResult_content']/div[2]/div[3]/pre/code";
        //*[@id="DatabaseTopic_DatabaseTopic_GetConversionRateResults_content"]/div[2]/div[3]/pre/code

        public GetConversionRateResult(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement inputID => Driver.FindElement(By.XPath(result));
        IWebElement un => Driver.FindElement(By.XPath(uname));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void ConversionRateResult()
        {

            test = ExtReport.CreateTest("TC_028_GetConversionRateResult").Info("Test Started");

            string inputD = InputExcelAPI.GetCellData("WebAPI", 23, 3);
            /*string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 27, 3);
            string val2 = ValidationExcelAPI.GetCellData("DatabaseTopic", 27, 4);*/

            SqlCommand cmd = new SqlCommand("select * from ConversionRate where TradeID = @TradeID", con);
            cmd.Parameters.AddWithValue("@TradeID", inputD);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string id;
            string conRate;
            id = Convert.ToString(reader[0]);
            conRate = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();

            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            Thread.Sleep(2000);*/

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "Get_ConversionRateResult selected");
            // Thread.Sleep(3000);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(result)));

            inputID.SendKeys(inputD);
            test.Log(Status.Info, "TradeID: "+inputD+ " entered");
            // Thread.Sleep(2000);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(uname)));


            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(lattise_Try)));

            button_lattise.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            // Thread.Sleep(5000);

            WebDriverWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(lattise_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_lattise);
            action.Perform();

            string resBody = body_lattise.Text;
            string resCode = res_Code.Text;


            test.Log(Status.Info, "Verifying Value....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "Get_ConversionRateResult Response is " + resCode);

                if (resBody.Contains(id))
                {
                    test.Log(Status.Pass, "Get_ConversionRateResult Response body contains ID " + id + " & ConversionRate " + conRate);
                    test.Log(Status.Pass, "Test is Pass");
                }

                else
                {
                    test.Log(Status.Fail, "Get_ConversionRateResult Response is Failed");
                    Assert.Fail("Get_ConversionRateResult Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "Get_ConversionRateResult Response is " + resCode);
                Assert.Fail("Get_ConversionRateResult Response is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
