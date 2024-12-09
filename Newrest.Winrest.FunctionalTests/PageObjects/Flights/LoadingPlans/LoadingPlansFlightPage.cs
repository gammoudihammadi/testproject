using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans
{
    public class LoadingPlansFlightPage : PageBase
    {

        public LoadingPlansFlightPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ____________________________________ Constantes ____________________________________________

        private const string FLIGHT_NUMBER = "//*[@id=\"tabContentFlightContainer\"]/div[1]/h1/span";
        private const string COUNTER = "//*[@id=\"tabContentItemContainer\"]/div/h1/span";
        private const string START_DATE = "start-date-picker";
        private const string END_DATE = "end-date-picker";
        private const string SHOW_FLIGHT = "//*[@id=\"ShowFlightsNotLinked\"]";
        private const string SEARCH = "FlightNumber";
        private const string RESET_FILTER = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string SELECT_ITEM = "check-flight";
        private const string BTN_DROP = "//*[@id=\"div-body\"]/div/div[1]/div/div/button";
        private const string LINKED_BTN = "link-flights";
        private const string LINKED_BTN_SAVE = "start-linking-flights";
        private const string UNLINKED_BTN = "unlink-flights";



        //______________________________________ Variables ____________________________________________

        [FindsBy(How = How.XPath, Using = FLIGHT_NUMBER)]
        private IWebElement _flightNumber;

        [FindsBy(How = How.XPath, Using = COUNTER)]
        private IWebElement _counter;

        [FindsBy(How = How.Id, Using = LINKED_BTN_SAVE)]
        private IWebElement _linkedSave;

        [FindsBy(How = How.Id, Using = START_DATE)]
        private IWebElement _startDate;

        [FindsBy(How = How.Id, Using = END_DATE)]
        private IWebElement _endDate;

        [FindsBy(How = How.Id, Using = SELECT_ITEM)]
        private IWebElement _selectItem;

        [FindsBy(How = How.Id, Using = SEARCH)]
        private IWebElement _search;

        [FindsBy(How = How.XPath, Using = SHOW_FLIGHT)]
        private IWebElement _showFlight;

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.XPath, Using = BTN_DROP)]
        private IWebElement _btnDrop;

        [FindsBy(How = How.Id, Using = LINKED_BTN)]
        private IWebElement _btnLink;

        [FindsBy(How = How.Id, Using = UNLINKED_BTN)]
        private IWebElement _btnunLink;

        public string GetFlightNumber()
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(FLIGHT_NUMBER)))
            {
                _flightNumber = WaitForElementExists(By.XPath(FLIGHT_NUMBER));
                return _flightNumber.Text;
            }
            else
            {
                return "0";
            }
        }

        public void SetStartDate(DateTime date)
        {
            _startDate = WaitForElementIsVisible(By.Id(START_DATE));
            _startDate.SetValue(ControlType.DateTime, date);
            _startDate.SendKeys(Keys.Tab);
            WaitForLoad();
        }

        public string GetFlightCount()
        {
            _counter = WaitForElementExists(By.XPath(COUNTER));
            return _counter.Text;
        }


        public enum FilterType
        {
            Search,
            StartDate,
            EndDate,
            Showflightsnotlinked,
            ShowLinkedFlights,
            ShowAll
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.Search:
                    _search = WaitForElementIsVisible(By.Id(SEARCH));
                    _search.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.StartDate:
                    _startDate = WaitForElementIsVisible(By.Id(START_DATE));
                    _startDate.SetValue(ControlType.DateTime, value);
                    _startDate.SendKeys(Keys.Tab);
                    break;
                case FilterType.EndDate:
                    _endDate = WaitForElementIsVisible(By.Id(END_DATE));
                    _endDate.SetValue(ControlType.DateTime, value);
                    _endDate.SendKeys(Keys.Tab);
                    break;
                case FilterType.Showflightsnotlinked:
                    _showFlight = WaitForElementExists(By.XPath("//label[text()='Show unlinked flights only']/..//input"));
                    _showFlight.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                case FilterType.ShowLinkedFlights:
                    _showFlight = WaitForElementExists(By.XPath("//label[text()='Show linked flights only']/..//input"));
                    _showFlight.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                case FilterType.ShowAll:
                    _showFlight = WaitForElementExists(By.XPath("//label[text()='Show all']/..//input"));
                    _showFlight.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
        }
        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(FilterType.EndDate, DateUtils.Now);
            }
        }

        public void LinkNewFlight()
        {
            WaitForLoad();
            _selectItem = WaitForElementExists(By.Id(SELECT_ITEM));
            _selectItem.SetValue(ControlType.CheckBox, true);
            WaitForLoad();

            _btnDrop = WaitForElementExists(By.XPath(BTN_DROP));
            _btnDrop.Click();
            WaitForLoad();

            _btnLink = WaitForElementExists(By.Id(LINKED_BTN));
            _btnLink.Click();
            WaitForLoad();

            _linkedSave = WaitForElementExists(By.Id(LINKED_BTN_SAVE));
            _linkedSave.Click();
            WaitForLoad();

        }
        public void UnLinkNewFlight()
        {
            WaitForLoad();
            _selectItem = WaitForElementExists(By.Id(SELECT_ITEM));
            _selectItem.SetValue(ControlType.CheckBox, true);
            WaitForLoad();

            _btnDrop = WaitForElementExists(By.XPath(BTN_DROP));
            _btnDrop.Click();
            WaitForLoad();

            _btnunLink = WaitForElementExists(By.Id(UNLINKED_BTN));
            _btnunLink.Click();
            WaitForLoad();

            _linkedSave = WaitForElementExists(By.Id(LINKED_BTN_SAVE));
            _linkedSave.Click();
            WaitForLoad();

        }

    }
}
