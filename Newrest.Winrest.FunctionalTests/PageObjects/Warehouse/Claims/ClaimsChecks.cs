using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims
{
    public class ClaimsChecks : PageBase
    {

        private const string CHECK_ALL_YES = "//*/input[@class='radioBtn']";
        private const string CHECK_ALL_NO = "//*/input[@class='radioBtnNo']";
        // Onglets
        private const string CLAIMS = "hrefTabContentItems";

        public ClaimsChecks(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public ClaimsItem ClickOnItems()
        {
            var _items = WaitForElementIsVisible(By.Id(CLAIMS));
            _items.Click();
            WaitForLoad();

            return new ClaimsItem(_webDriver, _testContext);

        }

        public void CheckAllYes()
        {
            var allYes = _webDriver.FindElements(By.XPath(CHECK_ALL_YES));
            foreach (var element in allYes)
            {
                Assert.IsTrue(element.Selected);
            }
        }

        public void CheckAllNo()
        {
            var allNo = _webDriver.FindElements(By.XPath(CHECK_ALL_NO));
            foreach (var element in allNo)
            {
                Assert.IsTrue(element.Selected);
            }
        }

        public void SetAllNo()
        {
            var allNo = _webDriver.FindElements(By.XPath(CHECK_ALL_NO));
            foreach (var element in allNo)
            {
                element.SetValue(PageBase.ControlType.RadioButton, true);
            }
            WaitForLoad();
            Thread.Sleep(2000);
            WaitForLoad();
        }
    }
}
