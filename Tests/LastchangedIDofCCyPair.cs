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
    class LastchangedIDofCCyPair
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();


        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetlastchangeIdOfCcyPair']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetlastchangeIdOfCcyPair_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetlastchangeIdOfCcyPair_content']/div[2]/div[4]/pre",
                        userName = "//*[@id='DatabaseTopic_DatabaseTopic_GetlastchangeIdOfCcyPair_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetlastchangeIdOfCcyPair_content']/div[2]/div[3]/pre/code";


        public LastchangedIDofCCyPair(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void GetLastchangedIDofCCyPair()
        {

            test = ExtReport.CreateTest("TS_069_GetLastchangedIDofCCyPair").Info("Test Started");

            //SqlCommand cmd = new SqlCommand("select max(Id), code from CcyPair", con);
            SqlCommand cmd = new SqlCommand("select * from CcyPairsChanges where Id=(select max(Id) from CcyPairsChanges)", con);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            string lastId;
            string code;

            reader.Read();

            lastId = Convert.ToString(reader[0]);
            code = Convert.ToString(reader[1]);

            reader.Close();
            con.Close();

            /* string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

             XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

             var sheet = wb.GetSheetAt(5);
             var row = sheet.GetRow(91);
             var value = row.GetCell(3).NumericCellValue.ToString();
 */
            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();
            test.Log(Status.Info, "LastchangedIDofCCyPair selected");

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


            test.Log(Status.Info, "Verifying CodeResult Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "LastchangedIDofCCyPair Response is " + resCode);

                if (resBody.Contains(lastId))
                {
                    test.Log(Status.Pass, "LastchangedIDofCCyPair Response body contains Last changed CCyPairID " +lastId+" "+code);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "LastchangedIDofCCyPair Response is Failed! , contains "+resBody);
                    Assert.Fail("LastchangedIDofCCyPair Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "LastchangedIDofCCyPair Response is " + resCode);
                Assert.Fail("LastchangedIDofCCyPair Response code is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
