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
    class ExecutingBrokerName
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        string username = User.getEncodedUserName();
        SqlConnection con = DBConnection.GetConnection();


        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetExecutingbrokerName']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetExecutingbrokerName_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetExecutingbrokerName_content']/div[2]/div[4]/pre",
                         userName = "//*[@id='DatabaseTopic_DatabaseTopic_GetExecutingbrokerName_content']/form/table/tbody/tr/td/input[@name='Username']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetExecutingbrokerName_content']/div[2]/div[3]/pre/code";


        public ExecutingBrokerName(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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

        public void GetExecutingBrokerName()
        {

            test = ExtReport.CreateTest("TS_064_GetExecutingBrokerName").Info("Test Started");

            SqlCommand cmd = new SqlCommand("select * from Broker where PostingID = (select max(PostingID) from Broker)", con);

            //cmd.Parameters.AddWithValue("@ID", val);

            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();

            reader.Read();

            string bname;
            string pId;

            bname = Convert.ToString(reader[1]).Trim();
            pId = Convert.ToString(reader[6]);

            reader.Close();
            con.Close();

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

            test.Log(Status.Info, "GetExecutingBrokerName selected");

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


            //ExtTest.Log(Status.Info, "Verifying Values....");


            test.Log(Status.Info, "Verifying TradeDetails Values....");

            if (resCode == "200")
            {
                test.Log(Status.Pass, "GetExecutingBrokerName Response is " + resCode);

                if (resBody.Contains(pId) && resBody.Contains(bname))
                {
                    test.Log(Status.Pass, "GetExecutingBrokerName Response body contains latest PostingID: " + pId+ " & Name "+bname);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else
                {
                    test.Log(Status.Fail, "GetExecutingBrokerName Response is fail");
                    Assert.Fail("GetExecutingBrokerName Response  is " + resCode);
                }
            }
            else
            {
                test.Log(Status.Fail, "GetExecutingBrokerName Response is " + resCode);
                Assert.Fail("GetExecutingBrokerName Response code is " + resCode);
            }
            /* if (resCode == "200")
             {
                 ExtTest.Log(Status.Pass, "GetExecutingBrokerName Response code is " + resCode);

                 int getrowcount = InputExcelAPI.GetLastRowIndex("DatabaseTopic");
                 int getcolumncount = InputExcelAPI.GetColumnCountByRow("DatabaseTopic", getrowcount);
                 // IDictionary<string, string> numberNames = new Dictionary<string, string>();



                 for (int rowval = 1; rowval <= getrowcount; rowval++)
                 {
                     // get all details of row
                     List<string> dataset = new List<string>();
                     for (int colval = 0; colval < getcolumncount; colval++)
                     {
                         dataset.Add(InputExcelAPI.GetCellData("DatabaseTopic", rowval, colval));
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
                         ExtTest.Log(Status.Pass, "GetExecutingBrokerName Response body contains ID " + dataset[0] + "Cutoff Time is" + dataset[1]);
                     }
                     else
                     {
                         ExtTest.Log(Status.Fail, "GetExecutingBrokerName Response is failed!! not contains " + dataset[0] + "& " + dataset[1]);
                         Assert.Fail("Response doesn't containg " + dataset[0]);
                     }
                 }
                 ExtTest.Log(Status.Pass, "Test is Pass");



             }
             else
             {
                 ExtTest.Log(Status.Fail, "Cutoff Response is " + resCode);
                 Assert.Fail("Response code " + resCode);

             }*/
            /*
            for (int i = 72; i < 75; i++)
                {

                    var row = sheet.GetRow(i);
                    var value = row.GetCell(3).StringCellValue;

                    if (resBody.Contains(value))
                    {
                        test.Log(Status.Pass, "GetExecutingBrokerName Response body contains " + value);
                    }
                    else
                    {
                        test.Log(Status.Fail, "GetExecutingBrokerName Response is " + value);
                    }

                }

                if (resCode == "200")
                {
                    test.Log(Status.Pass, "GetExecutingBrokerName Response is " + resCode);
                    test.Log(Status.Pass, "Test is Pass");

                    }
                    else
                {
                    test.Log(Status.Fail, "Cutoff Response is " + resCode);
                }
            */
            url_lattise.Click();
            //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
