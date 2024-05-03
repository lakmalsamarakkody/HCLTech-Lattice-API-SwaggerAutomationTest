using OpenQA.Selenium.Chrome;
using System;
using System.IO;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;

namespace SwaggerWebAPI.Libs
{
    class Driver
    {
        ChromeDriver chromeDriver;
        //ChromeOptions chromeOptions;
        //static string homePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));

        public Driver(string driver, string driverVersion)
        {
            // new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig(), driverVersion, WebDriverManager.Helpers.Architecture.X64); // automatically detects latest web driver
            //string chromeFolderName = string.Concat(driver,"_", driverVersion);
            //string chromePath = Path.Combine(homePath, @"Source\Drivers\", chromeFolderName);
            //chromeOptions = new ChromeOptions();
            //chromeOptions.BinaryLocation = string.Concat(chromePath,@"\chrome.exe");
            //chromeDriver = new ChromeDriver(driverPath, chromeOptions);
            //chromeDriver = new ChromeDriver(chromeOptions);

            new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig(), VersionResolveStrategy.MatchingBrowser);
            chromeDriver = new ChromeDriver();
        }
        public ChromeDriver GetDriver()
        {
            return chromeDriver;
        }
    }
}
