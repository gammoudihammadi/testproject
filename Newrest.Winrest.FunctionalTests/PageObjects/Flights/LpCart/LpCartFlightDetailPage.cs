using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart
{
    public class LpCartFlightDetailPage : PageBase
    {
        // ____________________________________ Constantes __________________________________________
        private const string FLIGHT_NAMES = "/html/body/div[3]/div/div[2]/div/div/div[2]/div/table/tbody/tr[*]/td[1]";
        private const string FLIGHT_DATES = "/html/body/div[3]/div/div[2]/div/div/div[2]/div/table/tbody/tr[*]/td[5]";

        //filter
        private const string SEARCH_FILTER = "SearchPattern";
        private const string SORT_BY_FILTER = "cbSortBy";
        private const string START_DATE_FILTER = "date-from-picker";
        private const string END_DATE_FILTER = "date-to-picker";
        private const string DATE = "//*[@id=\"services-table\"]/tbody/tr[*]/td[5]";
        private const string SITE_FROM = "//*[@id=\"services-table\"]/tbody/tr[*]/td[2]";
        private const string ETD = "//*[@id=\"services-table\"]/tbody/tr[*]/td[4]";
        private const string FLIGHTS_NAMES = "//*[@id=\"services-table\"]/tbody/tr[*]/td[1]";

        // __________________________________ Filtres ________________________________________

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _search;

        [FindsBy(How = How.Id, Using = SORT_BY_FILTER)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = START_DATE_FILTER)]
        private IWebElement _startDate;

        [FindsBy(How = How.Id, Using = END_DATE_FILTER)]
        private IWebElement _endDate;

        // ___________________________________  Méthodes ___________________________________________

        public LpCartFlightDetailPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }


        public enum FilterType
        {
            Search,
            SortBy,
            StartDate,
            EndDate,
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.Search:
                    _search = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _search.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORT_BY_FILTER));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;
                case FilterType.StartDate:
                    _startDate = WaitForElementIsVisible(By.Id(START_DATE_FILTER));
                    _startDate.SetValue(ControlType.DateTime, value);
                    _startDate.SendKeys(Keys.Tab);
                    break;
                case FilterType.EndDate:
                    _endDate = WaitForElementIsVisible(By.Id(END_DATE_FILTER));
                    _endDate.SetValue(ControlType.DateTime, value);
                    _endDate.SendKeys(Keys.Tab);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void CloseDatePicker()
        {
            _endDate = WaitForElementIsVisible(By.Id(END_DATE_FILTER));
            _endDate.SendKeys(Keys.Tab);
        }
        public bool IsFlightName(string flightName)
        {
            bool isGoodFlightName = true;

            var FlightNames = _webDriver.FindElements(By.XPath(FLIGHT_NAMES));

            foreach (var elm in FlightNames)
            {
                if (elm.Text != flightName)
                    return false;
            }
            return isGoodFlightName;
        }

        public List<string> GetFlightName()
        {
            List<string> values = new List<string>();

            var FlightNames = _webDriver.FindElements(By.XPath(FLIGHT_NAMES));

            foreach (var elm in FlightNames)
            {
                values.Add(elm.Text.Trim());
            }
            return values;
        }

        public List<string> GetFlightDates()
        {
            List<string> values = new List<string>();

            var FlightDates = _webDriver.FindElements(By.XPath(FLIGHT_DATES));

            foreach (var elm in FlightDates)
            {
                values.Add(elm.Text.Trim());
            }
            return values;
        }

        public bool IsSortedByDate()
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            var dateFormat = _startDate.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? ci = new CultureInfo("fr-FR") : ci = new CultureInfo("en-US");

            bool valueBool = true;
            DateTime ancientDate = DateTime.MinValue;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            WaitForLoad();

            var elements = _webDriver.FindElements(By.XPath(DATE));

            for (int i = 0; i < elements.Count; i++)
            {
                var element = elements[i];
                string dateText = element.Text;
                DateTime date = DateTime.Parse(dateText, ci);

                if (i == 0)
                {
                    ancientDate = date;
                }

                try
                {
                    if (DateTime.Compare(ancientDate, date) > 0)
                    { valueBool = false; }

                    ancientDate = date;
                }
                catch
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool IsSortedByFrom()
        {
            bool valueBool = true;
            var ancienSiteFrom = "";
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            var elements = _webDriver.FindElements(By.XPath(SITE_FROM));

            for (int i = 0; i < elements.Count; i++)
            {
                var element = elements[i];
                string siteFrom = element.Text;

                if (i == 0)
                    ancienSiteFrom = siteFrom;

                try
                {
                    if (String.Compare(ancienSiteFrom, siteFrom) > 0)
                    { valueBool = false; }

                    ancienSiteFrom = siteFrom;
                }
                catch
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool IsSortedByETD()
        {
            bool valueBool = true;
            var ancienETD = "";
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            var elements = _webDriver.FindElements(By.XPath(ETD));

            for (int i = 0; i < elements.Count; i++)
            {
                var element = elements[i];
                string etd = element.Text;

                if (i == 0)
                    ancienETD = etd;

                try
                {
                    if (String.Compare(ancienETD, etd) > 0)
                    { valueBool = false; }

                    ancienETD = etd;
                }
                catch
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool IsDateRespected(DateTime fromDate, DateTime toDate, string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var dates = _webDriver.FindElements(By.XPath(DATE));

            if (dates.Count == 0)
                return false;

            foreach(var date in dates)
            {

                DateTime dateResult = DateTime.Parse(date.Text,ci).Date;
                
                if (DateTime.Compare(dateResult, fromDate) < 0 || DateTime.Compare(dateResult, toDate) > 0 )
                {
                    return false;
                }

            }

            return true;
        }
        public List<string> GetFlightNames()
        {
            WaitPageLoading();
            var listElement = _webDriver.FindElement(By.XPath(FLIGHTS_NAMES));
            var ordersText = listElement.Text.Trim();
            var ordersList = ordersText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(order => order.Trim()).ToList();
            return ordersList;
        }
    }
}
