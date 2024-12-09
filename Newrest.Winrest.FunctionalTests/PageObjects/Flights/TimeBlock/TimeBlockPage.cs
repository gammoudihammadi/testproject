using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.TimeBlock
{
    public class TimeBlockPage : PageBase
    {
        public TimeBlockPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        private const string SITES_FILTER = "cbSortBy";
        private const string RESET_FILTER = "//*/a[text()='Reset Filter']";
        private const string SEARCH_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[2]/input";
        private const string EXTEND_MENU = "//*[@id=\"tabContentItemContainer\"]/div/div[2]/div/button";
        private const string SCHEDULE_BTN = "/html/body/div[2]/div/div[2]/div/div[2]/div/div/a[1]";

        private const string FIRST_WORKSHOP_DATE_FROM = "//*/tbody/tr[1]/td[9]/div[1]/div/div/input";
        private const string FIRST_WORKSHOP_TIME_FROM = "//*/tbody/tr[1]/td[9]/div[1]/input";
        private const string FIRST_WORKSHOP_DATE_TO = "//*/tbody/tr[1]/td[9]/div[2]/div/div/input";
        private const string FIRST_WORKSHOP_TIME_TO = "//*/tbody/tr[1]/td[9]/div[2]/input";

        private const string SECOND_WORKSHOP_DATE_FROM = "//*/tbody/tr[1]/td[10]/div[1]/div/div/input";
        private const string SECOND_WORKSHOP_TIME_FROM = "//*/tbody/tr[1]/td[10]/div[1]/input";
        private const string SECOND_WORKSHOP_DATE_TO = "//*/tbody/tr[1]/td[10]/div[2]/div/div/input";
        private const string SECOND_WORKSHOP_TIME_TO = "//*/tbody/tr[1]/td[10]/div[2]/input";


        private const string TIME_BEFORE_ETD_KTD = "/html/body/div[2]/div/div[2]/div/div[2]/div/div/a[1]";
        

        private const string DAY_FROM = "//*/tbody/tr/td[9]/div[1]/div/input";
        private const string TIME_FROM = "//*/tbody/tr/td[9]/div[1]/input";
        private const string DAY_TO = "//*/tbody/tr/td[9]/div[2]/div/input";
        private const string TIME_TO = "//*/tbody/tr/td[9]/div[2]/input";
        private const string WORKSHOP_FROM = "//virtual-scroller/div[2]/tbody/tr/td[10]/button[1]";
        private const string WORKSHOP_TO = "//virtual-scroller/div[2]/tbody/tr/td[10]/button[2]";
        private const string WORKSHOP_1 = "//*[@id=\"list-item-with-action\"]/table/thead/tr/th[11]";
        private const string WORKSHOP_2 = "//*[@id=\"list-item-with-action\"]/table/thead/tr/th[12]";
        private const string FLIGHTS_NAMES = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[3]";
        private const string TIME_INPUT = "workshopValue_InitialWorkshopDateStart";
        private const string TIME = "14:30:00";
        private const string DATE_TO = "ProdDateTo";



        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.XPath, Using = TIME_INPUT)]
        private IWebElement _timeInput;

        [FindsBy(How = How.XPath, Using = DATE_TO)]
        private IWebElement _dateTo;

        public enum FilterType
        {
            Site,
            Search,
            MajorFlightsOnly,
            HideCancelledFlights,
            HourETA,
            HourETD,
            DateFrom,
            DateTo,
            Customer,
            Workshops,
            WorkshopsUncheck,
            WorkshopsCheckAll
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Site:
                    var sites = WaitForElementIsVisible(By.Id(SITES_FILTER));
                    SelectElement select = new SelectElement(sites);
                    select.SelectByText((string)value);
                    break;
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.XPath(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    WaitPageLoading();
                    break;
                case FilterType.MajorFlightsOnly:
                    var _majorFlightsOnly = WaitForElementExists(By.Id("IsMajor"));
                    new Actions(_webDriver).MoveToElement(_majorFlightsOnly).Perform();
                    _majorFlightsOnly = WaitForElementExists(By.Id("IsMajor"));
                    _majorFlightsOnly.SetValue(ControlType.CheckBox, (bool)value);
                    break;
                case FilterType.HideCancelledFlights:
                    var _HideCancelledFlights = WaitForElementExists(By.Id("HideCancelled"));
                    new Actions(_webDriver).MoveToElement(_HideCancelledFlights).Perform();
                    _HideCancelledFlights = WaitForElementExists(By.Id("HideCancelled"));
                    _HideCancelledFlights.SetValue(ControlType.CheckBox, (bool)value);
                    break;
                case FilterType.HourETA:
                    var _hourETA = WaitForElementExists(By.Id("From"));
                    new Actions(_webDriver).MoveToElement(_hourETA).Perform();
                    _hourETA = WaitForElementIsVisible(By.Id("From"));
                    _hourETA.SetValue(ControlType.TextBox, value);
                    _hourETA.Click();
                    break;
                case FilterType.HourETD:
                    var _hourETD = WaitForElementExists(By.Id("To"));
                    new Actions(_webDriver).MoveToElement(_hourETD).Perform();
                    _hourETD = WaitForElementIsVisible(By.Id("To"));
                    _hourETD.SetValue(ControlType.TextBox, value);
                    _hourETD.Click();
                    break;
                case FilterType.DateFrom:
                    var _dateFrom = WaitForElementIsVisible(By.Id("ProdDateFrom"));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    WaitForLoad();
                    break;
                case FilterType.DateTo:
                    var _dateTo = WaitForElementIsVisible(By.Id("ProdDateTo"));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    WaitForLoad();
                    break;
                case FilterType.Customer:
                    if (value is bool)
                    {
                        ComboBoxOptions cbOpt = new ComboBoxOptions("SelectedCustomers_ms", null)
                        { ClickCheckAllAtStart = true };
                        ComboBoxSelectById(cbOpt);
                    }
                    else
                    {
                        ComboBoxSelectById(new ComboBoxOptions("SelectedCustomers_ms", (string)value));
                    }
                    break;
           
                case FilterType.WorkshopsUncheck:
                    var workshopsUncheck = WaitForElementIsVisible(By.Id("SelectedWorkshopsIds_ms"));
                    workshopsUncheck.Click();

                    var unselectAllWorkshop = WaitForElementIsVisible(By.XPath("/html/body/div[10]/div/ul/li[2]/a/span[2]"));
                    unselectAllWorkshop.Click();

                    workshopsUncheck.Click();
                    break;
               
                case FilterType.Workshops:
                //    var workshopsFilter = WaitForElementIsVisible(By.Id("SelectedWorkshopsIds_ms"));
                //    workshopsFilter.Click();

                //    unselectAllWorkshop = WaitForElementIsVisible(By.XPath("/html/body/div[10]/div/ul/li[2]/a/span[2]"));
                //    unselectAllWorkshop.Click();

                //    var searchWorkshops = WaitForElementIsVisible(By.XPath("/html/body/div[10]/div/div/label/input"));
                //    searchWorkshops.SetValue(ControlType.TextBox, value);
                //    WaitForLoad();

                //    var customerToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                //    customerToCheck.SetValue(ControlType.CheckBox, true);

                //    workshopsFilter.Click();
                    if (value is bool)
                    {
                        ComboBoxOptions cbOpt = new ComboBoxOptions("SelectedWorkshopsIds_ms", null)
                        { ClickCheckAllAtStart = true };
                        ComboBoxSelectById(cbOpt);
                    }
                    else
                    {
                        ComboBoxSelectById(new ComboBoxOptions("SelectedWorkshopsIds_ms", (string)value,false));
                    }
                    break;

                case FilterType.WorkshopsCheckAll:
                    workshopsUncheck = WaitForElementIsVisible(By.Id("SelectedWorkshopsIds_ms"));
                    workshopsUncheck.Click();

                    var selectAllCategory = WaitForElementIsVisible(By.XPath("/html/body/div[10]/div/ul/li[1]/a/span[2]"));
                    selectAllCategory.Click();

                    workshopsUncheck.Click();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilters()
        {
            _webDriver.Navigate().Refresh();
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(TimeBlockPage.FilterType.DateFrom, DateUtils.Now);
                Filter(TimeBlockPage.FilterType.DateTo, DateUtils.Now.AddDays(1));
            }
        }

        public void SetFirstMajorFlightsOnly(bool value)
        {
            var checkBoxMajor = WaitForElementExists(By.XPath("//*/input[contains(@class,'updateMajor')]"));
            checkBoxMajor.SetValue(PageBase.ControlType.CheckBox, value);
        }

        public string GetFirstFlightNumber(int numeroLigne = 1)
        {
            IWebElement confirmFlightNo;
            confirmFlightNo = WaitForElementIsVisible(By.XPath(string.Format("//*/tbody/tr[{0}]/td[3]", numeroLigne)));
            return confirmFlightNo.Text;
        }

        public bool VerifyFilterByFlightNumbre(string fnumber)
        {
            if (CheckTotalNumber() == 0) { return false; }
            for (int i = 1; i <= CheckTotalNumber(); i++)
            {
                if (GetFirstFlightNumber() != fnumber) { return false; }
            }
            return true;
        }

        public void Schedule()
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();

            var scheduleBtn = WaitForElementIsVisible(By.XPath(SCHEDULE_BTN));
            scheduleBtn.Click();
            
        }

        public void ResetAll()
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();

            var scheduleBtn = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div/div[2]/div/div/button"));
            scheduleBtn.Click();

        }

        public void TimeBeforeEtdKtd()
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();

            var timeBeforeEtdKtd = WaitForElementIsVisible(By.XPath(TIME_BEFORE_ETD_KTD));
            timeBeforeEtdKtd.Click();
        }

        public void SetFirstWorkshopParams(string dateFrom = null , string timeFrom = null
                                            ,string dateTo = null,string timeTo = null)
        {
            if(dateFrom != null)
            {
                IWebElement firstWorkshopDateFrom;
                
                firstWorkshopDateFrom = WaitForElementIsVisible(By.XPath("//*/tbody/tr/td[11]/div[1]/div/div/input"));
               
                firstWorkshopDateFrom.Clear();
                firstWorkshopDateFrom.SendKeys(dateFrom);
                firstWorkshopDateFrom.SendKeys(Keys.Tab);
                WaitForLoad();
            }
            if(timeFrom != null)
            {
                IWebElement firstWorkshopTimeFrom;
               
                firstWorkshopTimeFrom = WaitForElementExists(By.XPath("//*/tbody/tr/td[11]/div[1]/input"));
               
                firstWorkshopTimeFrom.SendKeys(timeFrom);
                WaitForLoad();
            }
            if(dateTo != null)
            {
                IWebElement firstWorkshopDateTo;
                
                firstWorkshopDateTo = WaitForElementIsVisible(By.XPath("//*/tbody/tr/td[11]/div[2]/div/div/input"));
                
                firstWorkshopDateTo.Clear();
                firstWorkshopDateTo.SendKeys(dateTo);
                firstWorkshopDateTo.SendKeys(Keys.Tab);
                WaitForLoad();
            }
            if(timeTo!= null)
            {
                IWebElement firstWorkshopTimeTo;
               
                firstWorkshopTimeTo = WaitForElementIsVisible(By.XPath("//*/tbody/tr/td[11]/div[2]/input"));
               
                firstWorkshopTimeTo.SendKeys(timeTo);
                WaitForLoad();
            }


        }

        public void SetSecondWorkshopParams(string dateFrom = null, string timeFrom = null
                                            , string dateTo = null, string timeTo = null)
        {
            if (dateFrom != null)
            {
                IWebElement secondWorkshopDateFrom;
               
                secondWorkshopDateFrom = WaitForElementIsVisible(By.XPath("//*/tbody/tr/td[12]/div[1]/div/div/input"));

                secondWorkshopDateFrom.Clear();
                secondWorkshopDateFrom.SendKeys(dateFrom);
                secondWorkshopDateFrom.SendKeys(Keys.Tab);
                WaitForLoad();
            }
            if (timeFrom != null)
            {
                IWebElement secondWorkshopTimeFrom;
              
                secondWorkshopTimeFrom = WaitForElementExists(By.XPath("//*/tbody/tr/td[12]/div[1]/input"));
               
                secondWorkshopTimeFrom.SendKeys(timeFrom);
                WaitForLoad();
            }
            if (dateTo != null)
            {
                IWebElement secondWorkshopDateTo;
               
                secondWorkshopDateTo = WaitForElementIsVisible(By.XPath("//*/tbody/tr/td[12]/div[2]/div/div/input"));
               
                secondWorkshopDateTo.Clear();
                secondWorkshopDateTo.SendKeys(dateTo);
                secondWorkshopDateTo.SendKeys(Keys.Tab);
                WaitForLoad();
            }
            if (timeTo != null)
            {
                IWebElement secondWorkshopTimeTo;
             
                    secondWorkshopTimeTo = WaitForElementIsVisible(By.XPath("//*/tbody/tr/td[12]/div[2]/input"));
              
                secondWorkshopTimeTo.SendKeys(timeTo);
                WaitForLoad();
            }
        }
        public bool VerifyWorkshopDaysAndTime()
        {
            IWebElement dayFrom;
            IWebElement dayTo;
            IWebElement timeFrom;
            IWebElement timeTo;
              
                dayFrom = WaitForElementIsVisible(By.XPath("//*/tbody/tr/td[11]/div[1]/div/input"));
                dayTo = WaitForElementIsVisible(By.XPath("//*/tbody/tr/td[11]/div[2]/div/input"));
                timeFrom = WaitForElementExists(By.XPath("//*/tbody/tr/td[11]/div[1]/input"));
                timeTo = WaitForElementIsVisible(By.XPath("//*/tbody/tr/td[11]/div[2]/input"));

            if (dayFrom.GetAttribute("value") == dayTo.GetAttribute("value") && timeFrom.GetAttribute("value") == "02:00:00"
                && timeTo.GetAttribute("value") == "04:00:00" && timeFrom.GetAttribute("class").Contains("edited-input")
                && timeTo.GetAttribute("class").Contains("edited-input"))
            {
                return true;
            }
            return false;
        }

        public bool VerifyNotLate()
        {
            IWebElement fromWorkshop;
            IWebElement toWorkshop;
            fromWorkshop = WaitForElementIsVisible(By.XPath(WORKSHOP_FROM));
            toWorkshop = WaitForElementIsVisible(By.XPath(WORKSHOP_TO));
            if (fromWorkshop.GetAttribute("class").Contains("late") &&
                toWorkshop.GetAttribute("class").Contains("late"))
            {
                return false;
            }
            return true;
        }
        public bool VerifyColumnsExist(string value)
        {
            IWebElement column1;
            IWebElement column2;
            column1 = WaitForElementExists(By.XPath(WORKSHOP_1));
            column2 = WaitForElementExists(By.XPath(WORKSHOP_2));
            if (value == column1.Text || value == column2.Text)
            {
                return true;
            }
            return false;
        }

        public BulkChangeModal BulkChange()
        {
            ShowExtendedMenu();

            var _bulkChangedBtn = WaitForElementIsVisible(By.XPath("//*/a[contains(@href,'BulkChange')]"));
            _bulkChangedBtn.Click();
            WaitForLoad();

            return new BulkChangeModal(_webDriver,_testContext);
        }
        public bool VerifyPageRefresh()
        {
            string navigationType = (string)((IJavaScriptExecutor)_webDriver).ExecuteScript("return window.performance.getEntriesByType('navigation')[0].type;");
            if (navigationType == "reload")
            {
                return false;
            }
            return true;
        }
        public string GetDayCorte()
        {
            var day = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div/div[3]/div/table/tbody/tr[1]/td[12]/div[1]/div/input"));
            return day.GetAttribute("value");
        }
        public string GetHourCorte()
        {
            var day = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div/div[3]/div/table/tbody/tr[1]/td[12]/div[1]/input"));
            return day.GetAttribute("value");
        }       
        public string GetDayCorteEnsamblaje()
        {
            var day = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div/div[3]/div/table/tbody/tr[1]/td[11]/div[1]/div/input"));
            return day.GetAttribute("value");
        }
        public string GetHourCorteEnsamblaje()
        {
            var day = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div/div[3]/div/table/tbody/tr[1]/td[11]/div[1]/input"));
            return day.GetAttribute("value");
        }
        public List<string> GetResultFilght()
        {
            var flightsNames = _webDriver.FindElements(By.XPath(FLIGHTS_NAMES));
            return flightsNames.Select(f => f.Text.Trim()).ToList();
        }
        public void AddTimeBlock()
        {
            _timeInput = _webDriver.FindElement(By.Id(TIME_INPUT));
            _timeInput.Clear();
            _timeInput.SendKeys(TIME);

        }

        public void CloseDatePicker()
        {
            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateTo.SendKeys(Keys.Tab);
        }
    }
}
