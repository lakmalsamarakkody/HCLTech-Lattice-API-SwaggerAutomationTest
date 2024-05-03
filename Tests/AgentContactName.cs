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
    class AgentContactName
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetAgentContactName']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetAgentContactName_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetAgentContactName_content']/div[2]/div[4]/pre",
                         agent = "//*[@id='DatabaseTopic_DatabaseTopic_GetAgentContactName_content']/form/table/tbody/tr/td/input[@name='agentId']",
                         userName = "//*[@id='DatabaseTopic_DatabaseTopic_GetAgentContactName_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetAgentContactName_content']/div[2]/div[3]/pre/code";


        public AgentContactName(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement id => Driver.FindElement(By.XPath(agent));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetAgentContactName()
        {

            test = ExtReport.CreateTest("TS_040_GetAgentContactName").Info("Test Started");

            //string value = InputExcelAPI.GetCellData("WebAPI", 31, 3);

            SqlCommand cmd = new SqlCommand("select * from Trader where ID = (select max(ID) from Trader)",con);
            //SqlCommand cmd = new SqlCommand("select * from Trader where ID = @ID", con);

            //cmd.Parameters.AddWithValue("@ID", value);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string tname;
            string Id;
            tname = Convert.ToString(reader[1]).Trim();
            Id = Convert.ToString(reader[0]);

            reader.Close();
            con.Close();
            //string val1 = ValidationExcelAPI.GetCellData("DatabaseTopic", 52, 3);
            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(51);
             var value = row.GetCell(3).StringCellValue;
             Thread.Sleep(2000);*/

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "AgentContactName selected");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(agent)));

            id.SendKeys(Id);
            test.Log(Status.Info, "Agent ID : "+Id+ " entered");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(userName)));

            string input = string.Format(username);
            un.SendKeys(input);

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(lattise_Try)));

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
                test.Log(Status.Pass, "AgentContactName Response is " + resCode);

                if (resBody.Contains(tname))
                {
                    test.Log(Status.Pass, "AgentContactName Response body contains Agent Name: " +tname+" & ID: " +Id);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "AgentContactName Response is fail");
                    Assert.Fail("AgentContactName Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "AgentContactName Response is " + resCode);
                Assert.Fail("AgentContactName Response is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
