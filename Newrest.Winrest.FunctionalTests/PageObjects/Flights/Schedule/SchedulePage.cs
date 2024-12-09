using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Globalization;
using System.Threading;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Schedule

{
    public class SchedulePage : PageBase
    {
        public SchedulePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________

        // General
        private const string EXTENDED_BTN = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/button";
        private const string UNFOLD_BTN = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[@class=\"btn unfold-all-btn unfold-all-btn-auto\"]";
        private const string FOLD_BTN = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[@class=\"btn unfold-all-btn unfold-all-btn-auto unfolded\"]";

        // Tableau
        private const string ARROW_BTN = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[1]/span";
        private const string TABLE_CONTENT = "//*[starts-with(@id,\"content_\")]";
        private const string DAY1_BUTTON = "0_0_delay_1";
        private const string DAY_BUTTON = "0_0_delay_0";
        private const string SERVICE_LABEL = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[5]/span";
        private const string SERVICE_NON_PRODUCED = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[7]/span";
        private const string CHANGE_PROD = "Items_Items_0__Details_0__IsProduced";
        private const string PRODUCED_TODAY = "//tr/td[contains(text(),'{0}')]//ancestor::div[@class='panel panel-default']//input[@value='0']";
        private const string PRODUCED_YESTERDAY = "//tr/td[contains(text(),'{0}')]//ancestor::div[@class='panel panel-default']//input[@value='1']";
        private const string SERVICEATD = "//div[@id='list-item-with-action']/div[2]/div[1]/div/div[2]/table/tbody/tr/td[6]/span";
        // Filtres
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string RESET_FILTER_PATCH = "//*[@id=\"item-filter-form\"]/div[1]/a";

        private const string SITE_FILTER = "cbSites";
        private const string FILTER_DATE_FROM = "ProdDateFrom";
        private const string FILTER_DATE_TO = "ProdDateTo";
        private const string CUSTOMER_FILTER = "SelectedCustomer_ms";
        private const string SEARCH_CUSTOMER = "/html/body/div[10]/div/div/label/input";
        private const string UNSELECT_ALL_CUSTOMERS = "/html/body/div[10]/div/ul/li[2]/a";

        private const string DATES = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string CUSTOMERS = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[3]";
        private const string FLIGHT_NUMBERS = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[4]";

        //__________________________________ Variables _____________________________________

        // General
        [FindsBy(How = How.XPath, Using = EXTENDED_BTN)]
        private IWebElement _extendedBtn;

        [FindsBy(How = How.XPath, Using = UNFOLD_BTN)]
        private IWebElement _unfold;

        [FindsBy(How = How.XPath, Using = FOLD_BTN)]
        private IWebElement _fold;

        // Tableau

        [FindsBy(How = How.XPath, Using = TABLE_CONTENT)]
        private IWebElement _tableContent;

        [FindsBy(How = How.Id, Using = DAY1_BUTTON)]
        private IWebElement _day1Btn;

        [FindsBy(How = How.Id, Using = DAY_BUTTON)]
        private IWebElement _dayBtn;

        [FindsBy(How = How.XPath, Using = SERVICE_LABEL)]
        private IWebElement _serviceLabel;

        [FindsBy(How = How.XPath, Using = SERVICE_NON_PRODUCED)]
        private IWebElement _serviceNonProduced;

        [FindsBy(How = How.Id, Using = CHANGE_PROD)]
        private IWebElement _changeProd;

        //__________________________________ Filtres _______________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;
        
        [FindsBy(How = How.XPath, Using = RESET_FILTER_PATCH)]
        private IWebElement _resetFilterPatch;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_CUSTOMERS)]
        private IWebElement _unselectAllCustomers;



        public enum FilterType
        {
            Site,
            DateFrom,
            DateTo,
            Customers
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Site:
                    _siteFilter = WaitForElementIsVisible(By.Id(SITE_FILTER));
                    _siteFilter.SetValue(ControlType.DropDownList, value);
                    break;
                case FilterType.DateFrom:
                    _dateFrom = WaitForElementIsVisible(By.Id(FILTER_DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;
                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisible(By.Id(FILTER_DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;
                case FilterType.Customers:
                    _customerFilter = WaitForElementIsVisible(By.Id(CUSTOMER_FILTER));
                    _customerFilter.Click();

                    _unselectAllCustomers = _webDriver.FindElement(By.XPath(UNSELECT_ALL_CUSTOMERS));
                    _unselectAllCustomers.Click();

                    _searchCustomer = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER));
                    _searchCustomer.SetValue(ControlType.TextBox, value);

                    var valueToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    valueToCheck.Click();

                    _customerFilter.Click();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilter()
        {
            if(isElementVisible(By.Id(RESET_FILTER_DEV)))
            {
                _resetFilterDev = WaitForElementIsVisible(By.Id(RESET_FILTER_DEV));
                _resetFilterDev.Click();
            }
            else
            {
                _resetFilterPatch = WaitForElementIsVisible(By.XPath(RESET_FILTER_PATCH));
                _resetFilterPatch.Click();
            }
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(FilterType.DateTo, DateUtils.Now);
            }
        }

        public bool IsFromToDateRespected(DateTime fromDate, DateTime toDate, string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var dates = _webDriver.FindElements(By.XPath(DATES));

            if (dates.Count == 0)
                return false;

            foreach (var elm in dates)
            {
                try
                {
                    DateTime date = DateTime.Parse(elm.Text, ci).Date;

                    if (DateTime.Compare(date, fromDate) < 0 || DateTime.Compare(date, toDate) > 0)
                        return false;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool VerifyCustomer(string customer)
        {
            var customers = _webDriver.FindElements(By.XPath(CUSTOMERS));

            if (customers.Count == 0)
                return false;

            foreach (var elm in customers)
            {
                if (!elm.Text.Contains(customer))
                {
                    return false;
                }
            }
            return true;
        }

        public bool IsFlightNumberPresent(string flightNumber)
        {
            var flightNumbers = _webDriver.FindElements(By.XPath(FLIGHT_NUMBERS));

            if (flightNumbers.Count == 0)
                return false;

            foreach (var elm in flightNumbers)
            {
                if (elm.Text.Equals(flightNumber))
                {
                    return true;
                }
            }
            return false;
        }

        // __________________________________ Méthodes _______________________________________

        // General
        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);

            _extendedBtn = WaitForElementIsVisible(By.XPath(EXTENDED_BTN));
            actions.MoveToElement(_extendedBtn).Perform();
        }

        public void UnfoldAll()
        {
            ShowExtendedMenu();
            WaitForLoad();
            _unfold = WaitForElementIsVisible(By.XPath(UNFOLD_BTN));
            _unfold.Click();
            // animation
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public void FoldAll()
        {
            ShowExtendedMenu();

            _fold = WaitForElementIsVisible(By.XPath(FOLD_BTN));
            _fold.Click();
            // animation
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public Boolean IsFoldAll()
        {
            _tableContent = WaitForElementExists(By.XPath(TABLE_CONTENT));
            WaitForLoad();

            if (_tableContent.GetAttribute("class") == "panel-collapse collapse")
                return true;

            return false;
        }

        // Tableau
        public void SetProdDayD()
        {
            _dayBtn = WaitForElementIsVisible(By.Id(DAY_BUTTON));
            _dayBtn.SetValue(ControlType.RadioButton, true);
            WaitForLoad();
        }
        public void SetProdDayD_1()
        {
            _dayBtn = WaitForElementIsVisible(By.Id(DAY1_BUTTON));
            _dayBtn.SetValue(ControlType.RadioButton, true);
            WaitForLoad();
        }

        public int GetNomnbreServiceAtDayD_1()
        {
            _serviceLabel = WaitForElementIsVisible(By.XPath(SERVICE_LABEL));
            return Convert.ToInt32(_serviceLabel.GetAttribute("innerText"));
        }

        public int GetServiceNonProducedValue()
        {
            _serviceNonProduced = WaitForElementIsVisible(By.XPath(SERVICE_NON_PRODUCED));
            return Convert.ToInt32(_serviceNonProduced.GetAttribute("innerText"));
        }

        public void ChangeIsProduced(bool isProduced)
        {
            WaitForLoad();
            _changeProd = WaitForElementExists(By.Id(CHANGE_PROD));
            _changeProd.SetValue(ControlType.CheckBox, isProduced);
            WaitForLoad();

            // Temps de prise en compte de la modification
            WaitPageLoading();
        }

        public void SetFlightProduced(string flightName, bool today)
        {
            if (today)
            {   WaitForLoad();
                //PRODUCED_TODAY
                var producedToday = WaitForElementIsVisible(By.XPath(String.Format(PRODUCED_TODAY, flightName)));
                producedToday.SetValue(ControlType.RadioButton, true);
                producedToday.Click();
            }
            else
            {
                var producedYesterday = WaitForElementIsVisible(By.XPath(String.Format(PRODUCED_YESTERDAY, flightName)));
                producedYesterday.Click();
            }

            // Temps de prise en compte de la modification
            WaitPageLoading();
        }

        public bool IsFlightProducedToday(string flightName)
        {
            var producedToday = WaitForElementIsVisible(By.XPath(String.Format(PRODUCED_TODAY, flightName)));
            return producedToday.GetAttribute("checked").Equals("true");
        }

        public bool isServiceAtDayD()
        {
            _day1Btn = _webDriver.FindElement(By.Id(DAY_BUTTON));
            bool iSchecked ;
            bool.TryParse(_day1Btn.GetAttribute("checked"),out iSchecked);
            return iSchecked;
        }
        public int GetNomnbreServiceAtDayD()
        {
            _serviceLabel = WaitForElementIsVisible(By.XPath(SERVICEATD));
            return Convert.ToInt32(_serviceLabel.GetAttribute("innerText"));
        }


    }
}