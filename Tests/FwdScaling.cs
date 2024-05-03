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
    class FwdScaling
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetFwdScaling']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetFwdScaling_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetFwdScaling_content']/div[2]/div[4]/pre",
                         ccy = "//*[@id='DatabaseTopic_DatabaseTopic_GetFwdScaling_content']/form/table/tbody/tr/td/input[@name='ccypair']",
                         userName = "//*[@id='DatabaseTopic_DatabaseTopic_GetFwdScaling_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetFwdScaling_content']/div[2]/div[3]/pre/code";


        public FwdScaling(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement un => Driver.FindElement(By.XPath(userName));
        IWebElement ccypair => Driver.FindElement(By.XPath(ccy));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetFwdScaling()
        {

            test = ExtReport.CreateTest("Get_GetFwdScaling").Info("Test Started");

            string value = InputExcelAPI.GetCellData("WebAPI", 32, 3);

            //string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 54, 3);

            //SqlCommand cmd = new SqlCommand("select * from CcyPair where ID = (select max(ID) from CcyPair)", con);
            SqlCommand cmd = new SqlCommand("select * from CcyPair where Code = @Code", con);

            cmd.Parameters.AddWithValue("@Code", value);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string code;
            string fwdScl;
            code = Convert.ToString(reader[1]).Trim();
            fwdScl = Convert.ToString(reader[2]);

            reader.Close();
            con.Close();
            // string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

 
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "GetFwdScaling selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(ccy)));

            ccypair.SendKeys(value);
            test.Log(Status.Info, "CcyPair : "+value+ " entered");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));

            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_Try)));

            button_lattise.Click();
            test.Log(Status.Info, "Try it Now Button Clicked");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lattise_body)));

            Actions action = new Actions(Driver);
            action.MoveToElement(body_lattise);
            action.Perform();

            string resBody = body_lattise.Text;
            string resCode = res_Code.Text;


            test.Log(Status.Info, "Verifying Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GetFwdScaling Response is " + resCode);

                if (resBody.Contains(fwdScl))
                {
                    test.Log(Status.Pass, "GetFwdScaling Response body contains CcyPair: " +code+" & it's FwdScaling: " + fwdScl);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetFwdScaling Response is fail");
                    Assert.Fail("GetFwdScaling Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "GetFwdScaling Response is " + resCode);
                Assert.Fail("GetFwdScaling Response is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
