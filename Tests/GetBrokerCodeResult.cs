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
    class GetBrokerCodeResult
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest test;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string lattise_url = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerBrokerCodeResults']/div[1]/h3/span[2]/a",
                         lattise_Try = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerBrokerCodeResults_content']/form/div[2]/input",
                         lattise_code = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerBrokerCodeResults_content']/div[2]/div[4]/pre",
                          lst = "//div[@id='DatabaseTopic_DatabaseTopic_GetBrokerBrokerCodeResults_content']/form/table/tbody/tr/td/textarea[@name='resultlst']",
                        lattise_body = "//*[@id='DatabaseTopic_DatabaseTopic_GetBrokerBrokerCodeResults_content']/div[2]/div[3]/pre/code";

        public GetBrokerCodeResult(IWebDriver driver, Dictionary<string, ExcelAPI> ExcelApiList, ExtentReports ExtReport)
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
        IWebElement inputList => Driver.FindElement(By.XPath(lst));
        IWebElement res_Code => Driver.FindElement(By.XPath(lattise_code));


        public void Get_BrokerCodeResult()
        {
            test = ExtReport.CreateTest("GetBrokerCodeResult").Info("Test Started");

            string val1 = InputExcelAPI.GetCellData("WebAPI", 18, 3);
            //string response = ValidationExcelAPI.GetCellData("DatabaseTopic", 28, 3);

            con.Open();
            //SqlCommand sqlCommand = new SqlCommand(@"select bc.Code
            //                                from Broker as b
            //                                join BrokerBrokerCode as bbc on b.ID = bbc.BrokerID
            //                                join BrokerCode as bc on bbc.BrokerCodeID = bc.ID", con);
            //SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            //sqlDataReader.Read();
            //string val1 = sqlDataReader[0].ToString();
            //sqlDataReader.Close();

            SqlCommand sqlCommand2 = new SqlCommand(@"select b.Name
                                                    from Broker as b
                                                    join BrokerBrokerCode as bbc on b.ID = bbc.BrokerID
                                                    join BrokerCode as bc on bbc.BrokerCodeID = bc.ID
                                                    where bc.Code = @code", con);
            sqlCommand2.Parameters.AddWithValue("@code", val1);
            SqlDataReader sqlDataReader2 = sqlCommand2.ExecuteReader();
            sqlDataReader2.Read();
            string response = sqlDataReader2[0].ToString().Trim();
            sqlDataReader2.Close();
            con.Close();

            WebDriverWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(lattise_url)));

            url_lattise.Click();

           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(2000);

            test.Log(Status.Info, "GetBrokerCodeResult selected");

            string input = string.Format(@"['{0}','{1}']", username, val1);
            inputList.SendKeys(input);
            //test.Log(Status.Info, "Code : 360JA entered");

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
                test.Log(Status.Pass, "GetBrokerCodeResult Response is " + resCode);

                if (resBody.Contains(response))
                {
                    test.Log(Status.Pass, "GetBrokerCodeResult Response contains "+ response);
                    test.Log(Status.Pass, "Test is Pass");
                }
                else if (resBody == "[]")
                {
                        test.Log(Status.Warning, "GetBrokerCodeResult Response doesn't have any results, please check given broker code has expected outcome as per your database.");
                        Assert.Warn("GetBrokerCodeResult Response is "+ resBody);
                }

                else
                {
                    test.Log(Status.Fail, "GetBrokerCodeResult Response is Failed!");
                    Assert.Fail("GetBrokerCodeResult Response is "+ resCode);
                }
            }
            else if(resCode=="401")
            {
                test.Log(Status.Warning, "GetBrokerCodeResult Response is " +resCode+" Unautherized!! ");
                Assert.Warn("GetBrokerCodeResult Response is " +resCode+ " Unautherized!!");
            }
            else
            {
                test.Log(Status.Fail, "GetBrokerCodeResult Response is " + resCode);
                Assert.Fail("GetBrokerCodeResult Response is " + resCode);
            }

            url_lattise.Click();
           // Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

        }
    }
}
