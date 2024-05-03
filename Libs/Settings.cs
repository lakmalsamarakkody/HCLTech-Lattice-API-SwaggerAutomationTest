using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace SwaggerWebAPI.Libs
{
    public class Settings
    {
        string InputExcelPath, ValidationExcelPath, OutputExcelPath;
        string DataSource, InitialCatalog, Username, Password;
        string WebDriver, WebDriverVersion, WebDriverHomeUrl;
        string homePath, configPath;
        string ExtentReportPath;

        public Settings()
        {
            homePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
            configPath = Path.Combine(homePath, @"config.xml");
            var xmlString = File.ReadAllText(configPath);
            var stringReader = new StringReader(xmlString);
            var dataSet = new DataSet();
            dataSet.ReadXml(stringReader);

            foreach (DataRow row in dataSet.Tables["AppSettings"].Rows)
            {
                WebDriver = row["WebDriver"].ToString();
                WebDriverVersion = row["WebDriverVersion"].ToString();
                WebDriverHomeUrl = row["WebDriverHomeUrl"].ToString();
            }

            foreach (DataRow row in dataSet.Tables["DatabaseSettings"].Rows)
            {
                DataSource = row["DataSource"].ToString();
                InitialCatalog = row["InitialCatalog"].ToString();
                Username = row["Username"].ToString();
                Password = row["Password"].ToString();
            }

            foreach (DataRow row in dataSet.Tables["ExcelSettings"].Rows)
            {
                InputExcelPath = row["InputExcelPath"].ToString();
                ValidationExcelPath = row["ValidationExcelPath"].ToString();
                OutputExcelPath = row["OutputExcelPath"].ToString();
            }

            foreach (DataRow row in dataSet.Tables["ReportSettings"].Rows)
            {
                ExtentReportPath = row["ExtentReportPath"].ToString();
            }
        }

        public string GetDriverType()
        {
            return WebDriver;
        }
        public string GetDriverVersion()
        {
            return WebDriverVersion;
        }
        public string GetDriverHomeUrl()
        {
            return WebDriverHomeUrl;
        }
        public string GetDataSource()
        {
            return DataSource;
        }
        public string GetInitialCatalog()
        {
            return InitialCatalog;
        }
        public string GetUsername()
        {
            return Username;
        }
        public string GetPassword()
        {
            return Password;
        }
        public string GetInputExcelPath()
        {
            if(InputExcelPath == "DEFAULT" || InputExcelPath == "default")
            {
                return Path.Combine(homePath, @"Source\Inputs.xlsx");
            }
            else
            {
                return Path.Combine(InputExcelPath);
            }
        }
        public string GetValidationExcelPath()
        {
            if (ValidationExcelPath == "DEFAULT" || ValidationExcelPath == "default")
            {
                return Path.Combine(homePath, @"Source\Validations.xlsx");
            }
            else
            {
                return Path.Combine(ValidationExcelPath);
            }
        }
        public string GetOutputExcelPath()
        {
            if (OutputExcelPath == "DEFAULT" || OutputExcelPath == "default")
            {
                return Path.Combine(homePath, @"Output\Outputs.xlsx");
            }
            else
            {
                return Path.Combine(OutputExcelPath);
            }
        }
        public string GetExtentReportPath()
        {
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string date = timestamp.Substring(0, 8);
            string time = timestamp.Substring(8, 6);
            string filename = string.Format(@"{0}_{1}\index.html",date, time);
            if (ExtentReportPath == "DEFAULT" || ExtentReportPath == "default")
            {
                return Path.Combine(homePath, @"Output\Extent Reports\", filename);
            }
            else
            {
                return Path.Combine(ExtentReportPath, filename);
            }
            
        }
    }
}
