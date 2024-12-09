using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Jobs
{
    [TestClass]
    public class JobsPage : PageBase
    {
        public JobsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        private const string IMPORTHISTORY = "//*[@id=\"nav-tab\"]/li[1]/a";
        private const string NAV_TO_CEGID = "//*[@id=\"nav-tab\"]/li[3]/a";
        
        [FindsBy(How = How.XPath, Using = NAV_TO_CEGID)]
        private IWebElement _cegidTab;
        [FindsBy(How = How.XPath, Using = IMPORTHISTORY)]
        private IWebElement _importhistory;
        public bool AccessPage()
        {
            if (isElementExists(By.XPath(IMPORTHISTORY)))
            {
                return true;
            }

            return false;
        }
        public CegidJobPage GoToCegidJobPage()
        {
            _cegidTab = WaitForElementIsVisible(By.XPath(NAV_TO_CEGID));
            _cegidTab.Click();
            WaitPageLoading();
            return new CegidJobPage(_webDriver, _testContext);
        }
    }
}
