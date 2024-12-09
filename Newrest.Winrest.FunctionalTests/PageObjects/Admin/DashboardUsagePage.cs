using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Linq;
using System.Security.Policy;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Admin
{
    public class DashboardUsagePage : PageBase
    {

        private const string SITE_FILTER = "SelectedSites_ms";
        private const string RESET_FILTER = "ResetFilter";


        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;
        public DashboardUsagePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public void Filter(FilterType filterType,object value, bool clearAllPreviousSelection = true, bool ignoreAutoSuggest = false)
        {
            WaitForLoad();
            switch (filterType)
            {
                case FilterType.From:
                    var inputFrom = WaitForElementIsVisible(By.Id("date-picker-start"));
                    inputFrom.Clear();
                    inputFrom.SendKeys(value.ToString());
                    inputFrom.SendKeys(Keys.Tab);
                    break;
                case FilterType.To:
                    var inputTo = WaitForElementIsVisible(By.Id("date-picker-end"));
                    inputTo.Clear();
                    inputTo.SendKeys(value.ToString());
                    inputTo.SendKeys(Keys.Tab);
                    break;
                case FilterType.Site:
                    if (value is bool)
                    {
                        ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER, null) { ClickCheckAllAtStart = true };
                        ComboBoxSelectById(cbOpt);
                    }
                    else
                    {
                        ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER, (string)value) { ClickUncheckAllAtStart = clearAllPreviousSelection };
                        ComboBoxSelectById(cbOpt);
                    }
                    break;
            }
        }
        public enum FilterType
        {
            From,
            To,
            Site,
            Module
        }
        public bool VerifySlowMovingPercentage(double percentage)
        {
            WaitPageLoading();
            var percentageDashboard = WaitForElementExists(By.XPath("//*[@id=\"contentContainer\"]/div/div[9]/table[2]/tbody/tr[4]/td[2]"));
            if (percentageDashboard.Text.Contains(percentage.ToString()))
                return true;
            return false;
        }
        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();

            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //pas de date
            }
        }
    }
}
