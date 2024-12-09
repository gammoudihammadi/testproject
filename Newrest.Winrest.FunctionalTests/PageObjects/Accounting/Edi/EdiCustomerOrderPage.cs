using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Edi
{
    public class EdiCustomerOrderPage : PageBase
    {
        private const string RESET_FILTER = "ResetFilter";

        private const string FILTER_SEARCH_CUSTOMER_CODE = "CustomerCode";
        private const string FILTER_SEARCH_CO_NUMBER = "CustomerOrderNumber";
        private const string FILTER_FROM = "date-picker-start";
        private const string FILTER_TO = "date-picker-end";
        private const string FILTER_SITES = "SelectedSites_ms";
        private const string FILTER_SITES_VALUE = "//*/button[@id='SelectedSites_ms']/span[2]";
        private const string FILTER_SHOW_ALL = "//*/input[@name='Status' and @value='-1']";
        private const string FILTER_IMPORT = "//*/input[@name='Status' and @value='1']";
        private const string FILTER_ERROR = "//*/input[@name='Status' and @value='2']";
        private const string CUSTOMER_ORDERS_DATES_LIST = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[5]";


        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_SEARCH_CUSTOMER_CODE)]
        private IWebElement _searchCustomerCodeFilter;
        [FindsBy(How = How.Id, Using = FILTER_SEARCH_CO_NUMBER)]
        private IWebElement _searchCONumberFilter;
        [FindsBy(How = How.Id, Using = FILTER_FROM)]
        private IWebElement _fromFilter;
        [FindsBy(How = How.Id, Using = FILTER_TO)]
        private IWebElement _toFilter;
        [FindsBy(How = How.Id, Using = FILTER_SITES)]
        private IWebElement _sitesFilter;
        [FindsBy(How = How.XPath, Using = FILTER_SHOW_ALL)]
        private IWebElement _showAllFilter;
        [FindsBy(How = How.XPath, Using = FILTER_IMPORT)]
        private IWebElement _importFilter;
        [FindsBy(How = How.XPath, Using = FILTER_ERROR)]
        private IWebElement _errorFilter;

        public enum FilterEdi
        {
            SearchCustomerCode,
            SearchCONumber,
            From,
            To,
            Sites,
            ShowAll,
            Import,
            Error
        }

        public EdiCustomerOrderPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
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

        public void Filter(FilterEdi FilterEdi, object value)
        {
            switch (FilterEdi)
            {
                case FilterEdi.SearchCustomerCode:
                    _searchCustomerCodeFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH_CUSTOMER_CODE));
                    _searchCustomerCodeFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterEdi.SearchCONumber:
                    _searchCONumberFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH_CO_NUMBER));
                    _searchCONumberFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterEdi.From:
                    _fromFilter = WaitForElementIsVisible(By.Id(FILTER_FROM));
                    _fromFilter.SetValue(ControlType.DateTime, value);
                    _fromFilter.SendKeys(Keys.Tab);
                    break;
                case FilterEdi.To:
                    _toFilter = WaitForElementIsVisible(By.Id(FILTER_TO));
                    _toFilter.SetValue(ControlType.DateTime, value);
                    _toFilter.SendKeys(Keys.Tab);
                    break;
                case FilterEdi.Sites:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITES, (string)value));
                    break;
                case FilterEdi.ShowAll:
                    _showAllFilter = WaitForElementExists(By.XPath(FILTER_SHOW_ALL));
                    _showAllFilter.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterEdi.Import:
                    _importFilter = WaitForElementExists(By.XPath(FILTER_IMPORT));
                    _importFilter.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterEdi.Error:
                    _errorFilter = WaitForElementExists(By.XPath(FILTER_ERROR));
                    _errorFilter.SetValue(ControlType.RadioButton, value);
                    break;
            }
            WaitForLoad();
        }


        public object GetFilterValue(FilterEdi FilterEdi)
        {
            switch (FilterEdi)
            {
                case FilterEdi.SearchCustomerCode:
                    _searchCustomerCodeFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH_CUSTOMER_CODE));
                    return _searchCustomerCodeFilter.GetAttribute("value");
                case FilterEdi.SearchCONumber:
                    _searchCONumberFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH_CO_NUMBER));
                    return _searchCONumberFilter.GetAttribute("value");
                case FilterEdi.From:
                    _fromFilter = WaitForElementIsVisible(By.Id(FILTER_FROM));
                    return _fromFilter.GetAttribute("value");
                case FilterEdi.To:
                    _toFilter = WaitForElementIsVisible(By.Id(FILTER_TO));
                    return _toFilter.GetAttribute("value");
                case FilterEdi.Sites:
                    _sitesFilter = WaitForElementIsVisible(By.XPath(FILTER_SITES_VALUE));
                    return _sitesFilter.Text;
                case FilterEdi.ShowAll:
                    _showAllFilter = WaitForElementIsVisible(By.XPath(FILTER_SHOW_ALL));
                    return _showAllFilter.Selected;
                case FilterEdi.Import:
                    _importFilter = WaitForElementIsVisible(By.XPath(FILTER_IMPORT));
                    return _importFilter.Selected;
                case FilterEdi.Error:
                    _errorFilter = WaitForElementIsVisible(By.XPath(FILTER_ERROR));
                    return _errorFilter.Selected;
            }
            return null;
        }
        public IEnumerable<string> GetDatesListCustomerOrders()
        {
            var listDatesPurchaseOrders = _webDriver.FindElements(By.XPath(CUSTOMER_ORDERS_DATES_LIST));
            return listDatesPurchaseOrders.Select(p => p.Text);
        }
        public bool VerifyDatesCustomerOrdersInFilter(IEnumerable<string> dates, DateTime from, DateTime to)
        {
            foreach (var date in dates)
            {
                if (DateTime.ParseExact(date, "dd/MM/yyyy", new CultureInfo("fr-FR")) > to || (DateTime.ParseExact(date, "dd/MM/yyyy", new CultureInfo("fr-FR")) < from))
                {
                    return false;
                }
            }
            return true;
        }
    }

}
