using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;

namespace Newrest.Winrest.FunctionalTests
{
    public class WebDriverFactory
    {
        private static IWebDriver _webDriver;

        public static IWebDriver Driver
        {
            get
            {
                if (_webDriver == null)
                    throw new NullReferenceException("The WebDriver browser instance was not initialized. You should first call the method InitBrowser.");
                return _webDriver;
            }
        }
        public static void EnsureBrowserClosed()
        {
            if (_webDriver != null)
            {
                CloseDriver();
            }
        }
        public static void InitializeDriver(string downloadsPath)
        {
            if (!Directory.Exists(downloadsPath))
                Directory.CreateDirectory(downloadsPath);

            string browser = "Chrome";
            switch (browser)
            {
                case "Chrome":
                    var options = new ChromeOptions();
                    options.AddUserProfilePreference("download.default_directory", downloadsPath);
                    options.AddUserProfilePreference("credentials_enable_service", false);
                    options.AddUserProfilePreference("profile.password_manager_enabled", false);
                    options.AddExcludedArgument("enable-automation");
                    options.AddAdditionalCapability("useAutomationExtension", false);
                    options.AddArgument("--lang=en-US");
                    options.AddArgument("--disable-search-engine-choice-screen");
                    _webDriver = new ChromeDriver(options);
                    break;
                case "Firefox":
                    _webDriver = new FirefoxDriver();
                    break;
                case "IE":
                    _webDriver = new InternetExplorerDriver();
                    break;
                default:
                    _webDriver = new ChromeDriver();
                    break;
            }
        }

        public static void CloseDriver()
        {
            if (_webDriver != null)
            {
                _webDriver.Close();
                _webDriver.Quit();
                _webDriver.Dispose();
                _webDriver = null;
            }
        }
    }
}
