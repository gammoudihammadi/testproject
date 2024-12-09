using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Jobs
{
    public class CegidJobPage : PageBase
    {
        private const string RESET_FILTER = "ResetFilter";
        private const string FIRST_ELEMENT = "//*[@id=\"tableListMenu\"]/tbody/tr[1]";

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How=How.XPath, Using =FIRST_ELEMENT)]
        private IWebElement _firstElement;
        public CegidJobPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
        }
        public bool HasJob()
        {
            return isElementExists(By.XPath(FIRST_ELEMENT));
        }
        public CegidJobDetailPage ClickFirstElement()
        {
            _firstElement = WaitForElementIsVisible(By.XPath(FIRST_ELEMENT));
            _firstElement.Click();
            WaitPageLoading();
            return new CegidJobDetailPage(_webDriver, _testContext);
        }

    }
}
