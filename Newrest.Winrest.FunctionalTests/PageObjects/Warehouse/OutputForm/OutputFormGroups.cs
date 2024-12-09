using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm
{
    public class OutputFormGroups : PageBase
    {
        public OutputFormGroups(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ___________________________________ Constantes ____________________________________________________

        private const string FILTER_GROUP = "SelectedGroups_ms";
        private const string UNCHECK_ALL_GROUPS = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_GROUP = "/html/body/div[11]/div/div/label/input";

        private const string GROUP_LINES = "//*[@id=\"detailedGroupsTable\"]/tbody/tr[*]";

        // ___________________________________ Variables _____________________________________________________

        [FindsBy(How = How.Id, Using = FILTER_GROUP)]
        private IWebElement _groupFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_GROUPS)]
        private IWebElement _uncheckAllGroups;

        [FindsBy(How = How.XPath, Using = SEARCH_GROUP)]
        private IWebElement _searchGroup;

        // ___________________________________ Méthodes ______________________________________________________

        public void FilterByGroup(string group)
        {
            _groupFilter = WaitForElementIsVisible(By.Id(FILTER_GROUP));
            _groupFilter.Click();

            _uncheckAllGroups = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_GROUPS));
            _uncheckAllGroups.Click();

            _searchGroup = WaitForElementIsVisible(By.XPath(SEARCH_GROUP));
            _searchGroup.SetValue(ControlType.TextBox, group);

            var valueToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + group + "']"));
            valueToCheck.Click();

            _groupFilter = WaitForElementIsVisible(By.Id(FILTER_GROUP));
            _groupFilter.Click();

            // lol
            WaitPageLoading();
            WaitPageLoading();
            WaitPageLoading();
        }

        public bool VerifyFilterGroup(string group)
        {
            var groups = _webDriver.FindElements(By.XPath(GROUP_LINES));

            if (groups.Count() < 2)
            {
                return false;
            }

            for (int i = 1; i < groups.Count; i++)
            {
                if (!groups[i].Text.Contains(group))
                {
                    return false;
                }
            }

            return true;
        }
    }
}
