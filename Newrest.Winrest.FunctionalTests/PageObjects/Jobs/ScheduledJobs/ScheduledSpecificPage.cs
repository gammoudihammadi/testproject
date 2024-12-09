using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs
{
    
    public class ScheduledSpecificPage : PageBase
    {
        private const string PLUS_BTN = "//*[@id=\"nav-tab\"]/li[5]/div/button";
        private const string NEW_SPECIFIC_JOB = "//*[@id=\"nav-tab\"]/li[5]/div/div/div/a";
        private const string FIRST_ROW_FREQUENCY = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[3]";
        private const string FIRST_JOB_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr/td[1]";



        // ----------- Filters -------------
        private const string SEARCH_BY_Name = "SearchPattern";
        private const string SHOW_ACTIVE_FILTER = "ShowOnlyActive";
        private const string SHOW_INACTIVE_ONLY = "ShowOnlyInactive";
        private const string SHOW_ALL = "ShowAll";
        private const string RESET_FILTER = "ResetFilter";
        private const string FIRST_LINE_NAME = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[2]";

        public ScheduledSpecificPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_____________________________________ Variables _____________________________________________
        // General
        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusbtn;

        [FindsBy(How = How.XPath, Using = NEW_SPECIFIC_JOB)]
        private IWebElement _newSpecific;

        [FindsBy(How = How.Id, Using = SEARCH_BY_Name)]
        private IWebElement _searchNameFilter;

        [FindsBy(How = How.Id, Using = SHOW_ACTIVE_FILTER)]
        private IWebElement _showActiveOnlyFilter;

        [FindsBy(How = How.Id, Using = SHOW_INACTIVE_ONLY)]
        private IWebElement _showInactiveOnlyFilter;

        [FindsBy(How = How.Id, Using = SHOW_ALL)]
        private IWebElement _showALLFilter;

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.XPath, Using = FIRST_ROW_FREQUENCY)]
        private IWebElement _firstRowFrequency;

        [FindsBy(How = How.XPath, Using = FIRST_JOB_ROW)]
        private IWebElement _firstJobRow;

        [FindsBy(How = How.XPath, Using = FIRST_LINE_NAME)]
        private IWebElement _firstLineName;

        public void ShowBtnPlus()
        {
            _plusbtn = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            _plusbtn.Click();
        }

        public enum FilterType
        {
            SearchByName,
            ShowAll,
            ShowActiveOnly,
            ShowInactiveOnly,

        }

        public void Filter(FilterType FilterType, object value)
        {
            switch (FilterType)
            {
                case FilterType.SearchByName:
                    _searchNameFilter = WaitForElementIsVisible(By.Id(SEARCH_BY_Name));
                    _searchNameFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    WaitPageLoading();
                    break;
                case FilterType.ShowActiveOnly:
                    _showActiveOnlyFilter = WaitForElementIsVisible(By.Id(SHOW_ACTIVE_FILTER));
                    _showActiveOnlyFilter.SetValue(ControlType.CheckBox, value);
                    WaitPageLoading();
                    break;
                case FilterType.ShowAll:
                    _showALLFilter = WaitForElementExists(By.Id(SHOW_ALL));
                    _showALLFilter.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;
                case FilterType.ShowInactiveOnly:
                    _showInactiveOnlyFilter = WaitForElementExists(By.Id(SHOW_INACTIVE_ONLY));
                    _showInactiveOnlyFilter.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;
            }
            WaitPageLoading();
            WaitForLoad();
        }

        public SpecificCreateModalPage OpenModalCreateSpecificJob()
        {
            ShowBtnPlus();
            _newSpecific = WaitForElementIsVisible(By.XPath(NEW_SPECIFIC_JOB));
            _newSpecific.Click();
            return new SpecificCreateModalPage(_webDriver, _testContext);
        }
        public int CountSpecificJobs()
        {
            var specificJob = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[2]"));
            var count = specificJob.Count();
            return count;
        }

        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
        }

        public string GetFirstLineName()
        {
            _firstLineName = WaitForElementExists(By.XPath(FIRST_LINE_NAME));
            WaitForLoad();
            return _firstLineName.Text;
        }

        public string GetFirstFrequency()
        {
            _firstRowFrequency = WaitForElementExists(By.XPath(FIRST_ROW_FREQUENCY));
            return _firstRowFrequency.Text;
        }

        public SpecificEditModal EditSpecific()
        {
            _firstJobRow = WaitForElementIsVisible(By.XPath(FIRST_JOB_ROW));
            _firstJobRow.Click();
            WaitForLoad();
            return new SpecificEditModal(_webDriver, _testContext);
        }
    }
}
