using Limilabs.Client.IMAP;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Trolleys;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight
{
    public class BarsetsPage : PageBase
    {
        public BarsetsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________

        // General
        private const string PLUS_BTN = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div[2]/div[2]/button";
        private const string NEW_BARSET_MULTIPLE = "btnAddMultipleFlight";
        private const string FLIGHTS_FILTER_ETD_FROM = "ETDFrom";
        private const string FLIGHTS_FILTER_ETD_TO = "ETDTo";
        private const string SEARCH_INPUT = "SearchPattern";
        private const string SITES = "cbSites";
        private const string STATUS_FILTER = "SelectedStatus_ms";
        private const string CUSTOMER_FILTER = "SelectedCustomerIds_ms";
        private const string EDIT_BTN = "//*[starts-with(@id,\"btnedit-\")]";
        private const string NUM_BARSET = "//*[@id=\"rowFlightBarset\"]/span/a";






        //__________________________________ Variables _____________________________________
        // General

        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusBtn;

        [FindsBy(How = How.Id, Using = NEW_BARSET_MULTIPLE)]
        private IWebElement _creatMultiBarset;

        [FindsBy(How = How.Id, Using = SITES)]
        private IWebElement _sites;

        [FindsBy(How = How.Id, Using = SEARCH_INPUT)]
        private IWebElement _searchInput;

        [FindsBy(How = How.Id, Using = FLIGHTS_FILTER_ETD_FROM)]
        private IWebElement _ETDFrom;

        [FindsBy(How = How.Id, Using = FLIGHTS_FILTER_ETD_TO)]
        private IWebElement _ETDTo;

        //___________________________________ Méthodes ____________________________________________

        // General
        public override void ShowPlusMenu()
        {
            _plusBtn = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            _plusBtn.Click();
            WaitForLoad();

        }

        public BarsetMultiCreateModalPage CreateMultiBarsets()
        {
            ShowPlusMenu();

            _creatMultiBarset = WaitForElementIsVisible(By.Id(NEW_BARSET_MULTIPLE));
            _creatMultiBarset.Click();
            WaitForLoad();

            return new BarsetMultiCreateModalPage(_webDriver, _testContext);
        }

        public enum FilterType
        {
            SearchFlight,
            Sites,
            Customer,
            Status,
            ETDFrom,
            ETDTo,
        }
        public void Filter(FilterType filterType, object value)
        {
            WaitForLoad();
            switch (filterType)
            {
                case FilterType.SearchFlight:
                    _searchInput = WaitForElementIsVisible(By.Id(SEARCH_INPUT));
                    _searchInput.SetValue(ControlType.TextBox, value);
                    // pourquoi ?
                    WaitPageLoading();
                    WaitForLoad();
                    break;
                case FilterType.Sites:
                    WaitForLoad();
                    _sites = WaitForElementIsVisible(By.Id(SITES));
                    _sites.SetValue(ControlType.DropDownList, value);
                    WaitForLoad();
                    break;
                case FilterType.Status:
                    string COLLAPSE_STATUS = "//*/label[text()='Status']/../../a[contains(@class,'collapsed')]";
                    if (isElementExists(By.XPath(COLLAPSE_STATUS)))
                    {
                        IWebElement collapse = WaitForElementExists(By.XPath(COLLAPSE_STATUS));
                        collapse.Click();
                        Thread.Sleep(2000);
                        WaitForLoad();
                    }
                    ComboBoxSelectById(new ComboBoxOptions(STATUS_FILTER, (string)value, true));
                    // pourquoi ?
                    WaitPageLoading();
                    break;
                case FilterType.Customer:
                    string COLLAPSE_CUSTOMER = "//*/label[text()='Customers']/../../a[contains(@class,'collapsed')]";
                    if (isElementVisible(By.XPath(COLLAPSE_CUSTOMER)))
                    {
                        IWebElement collapse = WaitForElementIsVisible(By.XPath(COLLAPSE_CUSTOMER));
                        collapse.Click();
                        Thread.Sleep(2000);
                        WaitForLoad();
                    }
                    ComboBoxSelectById(new ComboBoxOptions(CUSTOMER_FILTER, (string)value));
                    break;

                case FilterType.ETDFrom:
                    _ETDFrom = WaitForElementIsVisible(By.Id(FLIGHTS_FILTER_ETD_FROM));
                    _ETDFrom.Clear();
                    _ETDFrom.Click();
                    _ETDFrom.SendKeys((string)value);
                    // pourquoi ?
                    WaitPageLoading();
                    WaitForLoad();
                    break;
                case FilterType.ETDTo:
                    _ETDTo = WaitForElementIsVisible(By.Id(FLIGHTS_FILTER_ETD_TO));
                    _ETDTo.Clear();
                    _ETDTo.Click();
                    _ETDTo.SendKeys((string)value);
                    WaitPageLoading();
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }
            WaitPageLoading();
        }

        public BarsetEditModalPage EditFirstBarset()
        {
            WaitPageLoading();
            var _editBarsetBtn = WaitForElementIsVisible(By.XPath(EDIT_BTN));
            new Actions(_webDriver).MoveToElement(_editBarsetBtn).Click().Perform();
            WaitPageLoading();

            return new BarsetEditModalPage(_webDriver, _testContext);
        }

        public TrolleyAdjustingPage ClickOnNumBarset()
        {
            var _numBarset = WaitForElementIsVisible(By.XPath(NUM_BARSET));
            new Actions(_webDriver).MoveToElement(_numBarset).Click().Perform();
            WaitLoading();

            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new TrolleyAdjustingPage(_webDriver, _testContext);
        }
    }
}
