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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans
{
    public class LoadingPlansPage : PageBase
    {

        public LoadingPlansPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________ Constantes _____________________________________

        private const string PLUS_BTN = "//div[@class=\"dropdown dropdown-add-button\"]";

        private const string NEW_LOADING_PLAN = "//a[text()=\"New loading plan\"]";
        private const string FIRST_LP = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table";
        private const string END_DATE_FIRST_LP = "//*[@id=\"list-item-with-action\"]/div[2]/div/div/div[2]/table/tbody/tr/td[10]";
        private const string FIRST_LP_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div/div/div[2]/table/tbody/tr/td[2]";
        private const string FIRST_LP_TYPE = "//*[@id=\"list-item-with-action\"]/div[2]/div/div/div[2]/table/tbody/tr/td[3]";
        private const string DATE_FROM = "//*[@id=\"list-item-with-action\"]/div[*]/div/div/div[2]/table/tbody/tr/td[9]";
        private const string DATE_TO = "//*[@id=\"list-item-with-action\"]/div[*]/div/div/div[2]/table/tbody/tr/td[10]";
        private const string VERIFY_CUSTOMER = "//*[@id=\"list-item-with-action\"]/div[*]/div/div/div[2]/table/tbody/tr/td[4]";
        private const string ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string ITEM_ETD = "//*[@id=\"list-item-with-action\"]/div[*]/div/div/div[2]/table/tbody/tr/td[7]";
        
        private const string SEARCH_FILTER = "SearchPattern";
        private const string SORT_BY_FILTER = "cbSortBy";
        private const string SITES_FILTER = "cbSites";
        private const string AIRPORT_FILTER = "cbAirports";
        private const string START_DATE_FILTER = "start-date-picker";
        private const string END_DATE_FILTER = "end-date-picker";
        private const string FROM_HEURES_FILTER = "//*[@id=\"item-filter-form\"]/div[8]/div[1]/input[2]";
        private const string FROM_MINUTES_FILTER = "//*[@id=\"item-filter-form\"]/div[8]/div[1]/input[3]";
        private const string TO_HEURES_FILTER = "//*[@id=\"item-filter-form\"]/div[8]/div[2]/input[2]";
        private const string TO_MINUTES_FILTER = "//*[@id=\"item-filter-form\"]/div[8]/div[2]/input[3]";
        private const string ROUTE_FILTER = "Route";
        private const string EXPIRED_SERVICES_FILTER = "ShowLPWithExpiredServices";
        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string UNSELECT_CUSTOMERS = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_CUSTOMER = "/html/body/div[10]/div/div/label/input";

        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string RESET_FILTER_PATCH = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string MASSIVE_DELETE = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div[1]/div/a";
        private const string SEARCH_BTN = "SearchLoadingPlansBtn";
        private const string VERIFY_NAME = "//*[@id=\"list-item-with-action\"]/div[*]/div/div/div[2]/table/tbody/tr/td[2]";


        private const string FIRST_LOADING_PLAN_CUSTOMER = "/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/table/tbody/tr/td[4]";
        
        
        private const string LOADING_PLAN_FROM_DATES = "/html/body/div[2]/div/div[2]/div/div[2]/div[@class=\"panel panel-default\"]//tr[*]/td[9]";
        private const string LOADING_PLAN_TO_DATES = "/html/body/div[2]/div/div[2]/div/div[2]/div[@class=\"panel panel-default\"]//tr[*]/td[10]";
        private const string CONSULTER_DETAILS = "//*[@id=\"hrefTabContentDetails\"]";
        private const string FIRST_LOADINGPLAN = "//*[@id='addMultipleFlightForm']/div[2]/div/div/div[5]/div[1]/table/tbody/tr";

        // __________________________________ Variables ______________________________________

        [FindsBy(How = How.XPath, Using = NEW_LOADING_PLAN)]
        private IWebElement _createNewLoadingPlan;



        [FindsBy(How = How.XPath, Using = CONSULTER_DETAILS)]
        private IWebElement _ConsulterDetails;
        [FindsBy(How = How.XPath, Using = FIRST_LP)]
        private IWebElement _firstLoadingPlan;

        [FindsBy(How = How.XPath, Using = END_DATE_FIRST_LP)]
        private IWebElement _endDateFirstLP;

        [FindsBy(How = How.XPath, Using = FIRST_LP_NAME)]
        private IWebElement _firstLoadingPlanName;

        [FindsBy(How = How.XPath, Using = FIRST_LP_TYPE)]
        private IWebElement _firstLoadingPlanType;

        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusBtn;


        [FindsBy(How = How.XPath, Using = MASSIVE_DELETE)]
        private IWebElement _massiveDelete;
        // __________________________________ Filtres ________________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;
        
        [FindsBy(How = How.XPath, Using = RESET_FILTER_PATCH)]
        private IWebElement _resetFilterPatch;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _search;

        [FindsBy(How = How.Id, Using = SORT_BY_FILTER)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = SITES_FILTER)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = AIRPORT_FILTER)]
        private IWebElement _airportInput;

        [FindsBy(How = How.Id, Using = START_DATE_FILTER)]
        private IWebElement _startDate;

        [FindsBy(How = How.Id, Using = END_DATE_FILTER)]
        private IWebElement _endDate;

        [FindsBy(How = How.XPath, Using = FROM_HEURES_FILTER)]
        private IWebElement _fromHeures;

        [FindsBy(How = How.XPath, Using = FROM_MINUTES_FILTER)]
        private IWebElement _fromMin;

        [FindsBy(How = How.XPath, Using = TO_HEURES_FILTER)]
        private IWebElement _toHeures;

        [FindsBy(How = How.XPath, Using = TO_MINUTES_FILTER)]
        private IWebElement _toMin;

        [FindsBy(How = How.Id, Using = ROUTE_FILTER)]
        private IWebElement _routeInput;

        [FindsBy(How = How.Id, Using = EXPIRED_SERVICES_FILTER)]
        private IWebElement _expiredServiceInput;

        [FindsBy(How = How.Id, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_CUSTOMERS)]
        private IWebElement _unselectAllCustomers;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.XPath, Using = FIRST_LOADINGPLAN)]
        private IWebElement _firstloading;
        public enum FilterType
        {
            Search,
            SortBy,
            Site,
            Customers,
            Route,
            ExpiredService,
            AirportDeparture,
            StartDate,
            EndDate,
            From,
            To
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
                case FilterType.Site:
                    _site = WaitForElementIsVisible(By.Id(SITES_FILTER));
                    _site.SetValue(ControlType.DropDownList, value);
                    break;
                case FilterType.AirportDeparture:
                    _airportInput = WaitForElementIsVisible(By.Id(AIRPORT_FILTER));
                    _airportInput.SetValue(ControlType.DropDownList, value);
                    break;
                case FilterType.StartDate:
                    _startDate = WaitForElementIsVisible(By.Id(START_DATE_FILTER));
                    _startDate.SetValue(ControlType.DateTime, value);
                    _startDate.SendKeys(Keys.Tab);
                    _endDate = WaitForElementIsVisible(By.Id(END_DATE_FILTER));
                    _endDate.SendKeys(Keys.Tab);
                    break;
                case FilterType.EndDate:
                    _endDate = WaitForElementIsVisible(By.Id(END_DATE_FILTER));
                    _endDate.SetValue(ControlType.DateTime, value);
                    _endDate.SendKeys(Keys.Tab);
                    break;
                case FilterType.From:
                    string[] valuesFrom = null;
                    string stringValue = value as string;
                    if (stringValue != null)
                    {
                        valuesFrom = stringValue.Split(':');
                    }
                    if (valuesFrom.Length >= 2) {
                        _fromHeures = WaitForElementIsVisible(By.XPath(FROM_HEURES_FILTER));
                    _fromHeures.SetValue(ControlType.TextBox, valuesFrom[0]);

                    _fromMin = WaitForElementIsVisible(By.XPath(FROM_MINUTES_FILTER));
                    _fromMin.SetValue(ControlType.TextBox, valuesFrom[1]);
                    }
                    break;
                case FilterType.To:
                    string[] valuesTo=null;
                     stringValue = value as string;

                    if (stringValue != null)
                    {
                        valuesTo = stringValue.Split(':'); 
                    }
                    if (valuesTo.Length >=2) 
                    {
                        _toHeures = WaitForElementIsVisible(By.XPath(TO_HEURES_FILTER));
                        _toHeures.SetValue(ControlType.TextBox, valuesTo[0]);

                        _toMin = WaitForElementIsVisible(By.XPath(TO_MINUTES_FILTER));
                        _toMin.SetValue(ControlType.TextBox, valuesTo[1]);
                    }
                    break;
                case FilterType.Route:
                    _routeInput = WaitForElementIsVisible(By.Id(ROUTE_FILTER));
                    _routeInput.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.ExpiredService:
                    _expiredServiceInput = WaitForElementExists(By.Id(EXPIRED_SERVICES_FILTER));
                    action.MoveToElement(_expiredServiceInput).Perform();
                    _expiredServiceInput.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.Customers:
                    _customerFilter = WaitForElementExists(By.Id(CUSTOMER_FILTER));
                    action.MoveToElement(_customerFilter).Perform();
                    _customerFilter.Click();

                    // On décoche toutes les options
                    _unselectAllCustomers = _webDriver.FindElement(By.XPath(UNSELECT_CUSTOMERS));
                    _unselectAllCustomers.Click();

                    _searchCustomer = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER));
                    _searchCustomer.SetValue(ControlType.TextBox, value);

                    var valueToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    valueToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();

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
                // FilterType.To String[] de heure puis minute, pas de date ici
            }
        }

        public bool IsSortedByEtd()
        {
            bool valueBool = true;
            var ancienEtd = "";
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

            var elements = _webDriver.FindElements(By.XPath(ITEM_ETD));

            foreach (var element in elements)
            {
                try
                {
                    if (String.Compare(ancienEtd, element.Text) > 0)
                    { valueBool = false; }

                    ancienEtd = element.Text;
                }
                catch
                {
                    valueBool = false;
                }
            }



            return valueBool;
        }

        public bool IsSortedByDate()
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            var dateFormat = _startDate.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

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

            var elements = _webDriver.FindElements(By.XPath(DATE_FROM));


            foreach (var element in elements)
            {
                string dateText = element.Text;
                DateTime date = DateTime.Parse(dateText, ci);

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

        public bool IsSortedByName()
        {
            bool valueBool = true;
            var ancienName = "";
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

            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));


            foreach (var element in elements)
            {

                try
                {
                    if (String.Compare(ancienName, element.Text) > 0)
                    { valueBool = false; }

                    ancienName = element.Text;
                }
                catch 
                {
                    valueBool = false;

                }

            }


            return valueBool;
        }

        public bool VerifyCustomer(string value)
        {
            bool valueBool = true;
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

            var elements = _webDriver.FindElements(By.XPath(VERIFY_CUSTOMER));


            foreach (var element in elements)
            {
                if (element.Text != value)
                {
                    valueBool = false;
                }
            }


            return valueBool;
        }

        public bool IsDateRespected(DateTime fromDateInput, DateTime toDateInput)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            var dateFormat = _startDate.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            bool valueBool = true;
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

            var elementsFrom = _webDriver.FindElements(By.XPath(DATE_FROM));
            var elementsTo = _webDriver.FindElements(By.XPath(DATE_TO));

            var elementsToAndFrom = elementsFrom.Zip(elementsTo, (f, t) => new { ElementFrom = f, ElementTo = t });

            foreach (var element in elementsToAndFrom)
            {
                string fromDateText = element.ElementFrom.Text;
                DateTime startDate = DateTime.Parse(fromDateText, ci);

                string toDateText = element.ElementTo.Text;
                DateTime endDate = DateTime.Parse(toDateText, ci);

                try
                {
                    if ((startDate <= fromDateInput && toDateInput <= endDate) ||
                        (startDate <= toDateInput && toDateInput <= endDate) ||
                        (startDate <= fromDateInput && fromDateInput <= endDate))
                    { valueBool = true; }
                    else valueBool = false;
                }
                catch
                {
                    valueBool = false;
                }
            }


            return valueBool;
        }

        //___________________________________ Méthodes ___________________________________________


        

        public LoadingPlansCreateModalPage LoadingPlansCreatePage()
        {
            var actions = new Actions(_webDriver);
            ShowPlusMenu();
            WaitForLoad();
            _createNewLoadingPlan = WaitForElementIsVisible(By.XPath(NEW_LOADING_PLAN));
            actions.MoveToElement(_createNewLoadingPlan).Perform();
            _createNewLoadingPlan.Click();
            WaitPageLoading();

            return new LoadingPlansCreateModalPage(_webDriver, _testContext);
        }

        public LoadingPlansGeneralInformationsPage ClickOnFirstLoadingPlan()
        {
            _firstLoadingPlan = WaitForElementIsVisible(By.XPath(FIRST_LP));
            _firstLoadingPlan.Click();
            WaitPageLoading();
            WaitForLoad();
            WaitPageLoading();
            return new LoadingPlansGeneralInformationsPage(_webDriver, _testContext);
        }

        public void ShowPlusMenu()
        {
            var actions = new Actions(_webDriver);

            _plusBtn = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            actions.MoveToElement(_plusBtn).Perform();
        }

        public string GetEndDate(string decimalSeparatorValue)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = decimalSeparatorValue.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _endDateFirstLP = WaitForElementIsVisible(By.XPath(END_DATE_FIRST_LP));
            DateTime date = DateTime.Parse(_endDateFirstLP.Text, ci);

            return date.ToShortDateString();
        }

        public string GetFirstLoadingPlanName()
        {
            _firstLoadingPlanName = WaitForElementIsVisible(By.XPath(FIRST_LP_NAME));
            return _firstLoadingPlanName.Text;
        }

        public string GetFirstLoadingPlanNameType()
        {
            _firstLoadingPlanType = WaitForElementIsVisible(By.XPath(FIRST_LP_TYPE));
            return _firstLoadingPlanType.Text;
        }
        public void MassiveDeleteLoadingPlan(string loadingPlanNumber, string site, string customer, DateTime date)
        {
            ShowExtendedMenu();
            
            WaitForLoad();
            _massiveDelete = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE));
            _massiveDelete.Click();
            WaitForLoad();
            var searchPattern = WaitForElementIsVisible(By.XPath("//*[@id=\"formMassiveDeleteLoadingPlan\"]/div/div[1]/div/input[@id=\"SearchPattern\"]"));
            searchPattern.SendKeys(loadingPlanNumber);
            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", site,false));
            ComboBoxSelectById(new ComboBoxOptions("SelectedCustomersToDelete_ms", customer,false));

            var toInput = WaitForElementIsVisible(By.Id("dateLoadingPlanEnd"));
            toInput.Clear();
            toInput.SetValue(ControlType.DateTime, date);

            var searchbtn = WaitForElementIsVisible(By.Id(SEARCH_BTN));
            searchbtn.Click();
            WaitForLoad();
            if (isElementVisible(By.XPath("//*[@id=\"tableLoadingPlans\"]/tbody/tr")))
            {
                var checkInput = WaitForElementIsVisible(By.XPath("//*[@id=\"tableLoadingPlans\"]/tbody/tr/td[1]"));
                checkInput.Click();
                WaitForLoad();
            }
            var deleteBtn = WaitForElementIsVisible(By.Id("deleteLoadingPlanBtn"));
            deleteBtn.Click();
            WaitForLoad();
            var deleteConfirmBtn = WaitForElementIsVisible(By.Id("dataConfirmOK"));
            deleteConfirmBtn.Click();
            WaitLoading();
            WaitForLoad();
            var okButton = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[3]/button"));
            okButton.Click();
            WaitForLoad();
        }

        public string GetETDFromLoadingPlanName()
        {
            var etdFrom = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div/div/div[2]/table/tbody/tr/td[7]"));

            return etdFrom.Text;
        }
        public string GetETDToLoadingPlanName()
        {
            var etdTo = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div/div/div[2]/table/tbody/tr/td[8]"));

            return etdTo.Text;
        }

        public bool VerifyLoadingPlan(string value)
        {
            bool valueBool = false;
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

            var elements = _webDriver.FindElements(By.XPath(VERIFY_NAME));


            foreach (var element in elements)
            {
                if (element.Text == value)
                {
                    valueBool = true;
                }
            }


            return valueBool;
        }
        public LoadingPlanMassiveDeleteModalPage ClickMassiveDelete()
        {
            ShowExtendedMenu();
            WaitForElementIsVisible(By.XPath(MASSIVE_DELETE)).Click();
            return new LoadingPlanMassiveDeleteModalPage(_webDriver, _testContext);
        }

        public void ClickOnDetails()
        {
            var _detailsclick = WaitForElementIsVisible(By.XPath(CONSULTER_DETAILS));
             _detailsclick.Click();
         }

    


        public bool VerifyFromTo(DateTime from, DateTime to)
        {
            var listFromDates = _webDriver.FindElements(By.XPath(LOADING_PLAN_FROM_DATES));
            var listTODates = _webDriver.FindElements(By.XPath(LOADING_PLAN_TO_DATES));

            List<DateTime> allDates = new List<DateTime>();

            foreach (var dateElement in listFromDates)
            {
                DateTime parsedDate;
                if (DateTime.TryParseExact(dateElement.Text.Substring(0, 10), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                {
                    if (!(parsedDate.Date <= to.Date))
                        return false;
                }
            }

            //q.EndDate >= startDate && q.StartDate <= endDate

            foreach (var dateElement in listTODates)
            {
                DateTime parsedDate;
                if (DateTime.TryParseExact(dateElement.Text.Substring(0, 10), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                {
                    if (!(parsedDate.Date >= from.Date))
                        return false;
                }
            }

            return true;
        }
        public LoadingPlansPage SelectFirstLoadingPlan()
        {
            _firstloading = WaitForElementIsVisible(By.XPath(FIRST_LOADINGPLAN));
            WaitForLoad();
            return new LoadingPlansPage(_webDriver, _testContext);

        }
        public bool IsTypeEqualInFirstLoadingPlan(string expectedType)
        {
            _firstLoadingPlan = WaitForElementIsVisible(By.XPath(FIRST_LP));
            var actualTypeElement = _firstLoadingPlan.FindElement(By.XPath(FIRST_LP));
            var actualType = actualTypeElement.Text.Trim();
            return actualType.Equals(expectedType, StringComparison.OrdinalIgnoreCase);
        }

    }
}
