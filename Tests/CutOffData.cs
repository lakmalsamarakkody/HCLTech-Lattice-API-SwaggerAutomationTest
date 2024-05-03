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
    class CutOffData
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetCutOffData']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetCutOffData_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetCutOffData_content']/div[2]/div[4]/pre",
                         inputD = "//*[@id='DatabaseTopic_DatabaseTopic_GetCutOffData_content']/form/table/tbody/tr/td/textarea[@name='lstresult']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetCutOffData_content']/div[2]/div[3]/pre/code";

        public CutOffData(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement result => Driver.FindElement(By.XPath(inputD));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));

        public void GetCutOffData()
        {
            test = ExtReport.CreateTest("Get_CutOffData").Info("Test Started");

            /*string value = InputExcelAPI.GetCellData("WebAPI", 33, 3);

            string val = ValidationExcelAPI.GetCellData("DatabaseTopic", 58, 3);*/

            SqlCommand cmd = new SqlCommand(@"select TOP(1) c.Id as CutOffId, d.Id as DealId, pc.Centre, pc.Time from Deal as d 
                                                join PrimeBroker as p on p.Id = d.PrimeBrokerID
                                                join PrimeBrokerCutoff as pc on p.Id = pc.PrimeBrokerID
                                                join Cutoff as c on pc.CutoffID = c.Id
                                                where pc.Centre != 'NULL' and pc.Time! = 'NULL' order by DealId desc", con
            );

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string CuffOfftID, dealId, center, time;
            CuffOfftID = Convert.ToString(reader[0]);
            dealId = Convert.ToString(reader[1]);
            center = Convert.ToString(reader[2]);
            time = Convert.ToString(reader[3]);

            reader.Close();
            con.Close();

            /*string path = @"D:\Users\j_gamage\source\repos\Swagger\Swagger\Test.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));

            var sheet = wb.GetSheetAt(5);
            var row = sheet.GetRow(32);
            //var value = row.GetCell(3).NumericCellValue.ToString();
            //var val2 = row.GetCell(4).StringCellValue;

            Thread.Sleep(2000);*/

          string  inputData = $"'{CuffOfftID}','{dealId}'";

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "CutOffData selected");

            string input = string.Format($"[{inputData},'{username}']");
            result.SendKeys(input);

            test.Log(Status.Info, inputData + "entered");

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_Try)));

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
                test.Log(Status.Pass, "CutOffData Response is " + resCode);

                if (resBody.Contains($"{time}~{center}"))
                {
                    test.Log(Status.Pass, "CutOffData Response contains " +resBody);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Warning, "CutOffData Response is Failed! Instead in contains " + resBody);
                    Assert.Fail("CutOffData Response is " + resBody);
                }
            }
            else
            {
                test.Log(Status.Fail, "CutOffData Response is " + resCode);
                Assert.Fail("CutOffData Response is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
