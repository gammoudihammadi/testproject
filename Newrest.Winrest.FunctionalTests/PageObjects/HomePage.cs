using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;

namespace Newrest.Winrest.FunctionalTests.PageObjects
{
    public class HomePage : PageBase
    {
        public HomePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void Navigate()
        {
            var url = _testContext.Properties["Winrest_URL"].ToString();
            _webDriver.Navigate().GoToUrl(url);
        }
    }
}