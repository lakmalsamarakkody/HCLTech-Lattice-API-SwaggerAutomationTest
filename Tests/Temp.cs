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
    class Temp
    {
        IWebDriver Driver;
        ExcelAPI InputExcelAPI, ValidationExcelAPI, OutputExcelAPI;
        ExtentReports ExtReport;
        ExtentTest ExtTest;
        WebDriverWait WebDriverWait;
        SqlConnection con = DBConnection.GetConnection();
        string username = User.getEncodedUserName();

        static string
       URL = "//*[@id='Connection_Connection_GetConnection']/div/h3/span/a";


        public Temp(IWebDriver driver, Dictionary<string,ExcelAPI> ExcelApiList, ExtentReports ExtReport)
        {
            this.Driver = driver;
            this.InputExcelAPI = ExcelApiList["InputExcelApi"];
            this.ValidationExcelAPI = ExcelApiList["ValidationExcelApi"];
            this.OutputExcelAPI = ExcelApiList["OutputExcelApi"];
            this.ExtReport = ExtReport;
            this.WebDriverWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
        }

        IWebElement conn_url => Driver.FindElement(By.XPath(URL));


        public void runTemp()
        {
            string cellValue = ValidationExcelAPI.GetCellData(1, 0);
            string cellValue1 = ValidationExcelAPI.GetCellData(2, 1, 0);
            string cellValue2 = ValidationExcelAPI.GetCellData("Cutoffs", 1, 0);
            string cellValue3 = InputExcelAPI.GetCellData("Cutoffs", 1, 0);
            string cellValue4 = OutputExcelAPI.GetCellData("Cutoffs", 1, 0);
            int lastRowIndex = ValidationExcelAPI.GetLastRowIndex("Cutoffs");
            int rowCount = ValidationExcelAPI.GetRowCount("Cutoffs");
            int colCountHeader = ValidationExcelAPI.GetColumnCountByHeader("Cutoffs");
            int colCount = ValidationExcelAPI.GetColumnCountByRow("Cutoffs", 0);
            bool value = ValidationExcelAPI.SetCellData("value");
            bool value1 = ValidationExcelAPI.SetCellData(2, 2, "value");
            bool value2 = ValidationExcelAPI.SetCellData(4, 5, 7, "test value");
            bool value3 = ValidationExcelAPI.SetCellData("Cutoffs", 3, 2, "test value");

 
        }
    }
}
