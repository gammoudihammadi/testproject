using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Commons.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus
{
    public class MenusPage : PageBase
    {
        // ______________________________________ Constantes _____________________________________________

        // Général
        private const string PLUS_BUTTON = "//*/button[text()='+']";
        private const string NEW_MENU = "New Menu";

        // Tableau menus
        private const string FIRST_MENU_NAME = "//*[@id=\"tableListMenu\"]/tbody/tr/td[3]";
        private const string FIRST_MENU_SITE = "//*[@id=\"tableListMenu\"]/tbody/tr/td[4]";
        private const string SECOND_MENU_NAME = "//*[@id=\"tableListMenu\"]/tbody/tr[2]/td[3]";

        // Filtres
        private const string RESET_FILTER = "//*/a[text()='Reset Filter']";

        private const string SEARCH_FILTER = "SearchPattern";
        private const string SORTBY_FILTER = "cbSortBy";

        private const string SHOW_ALL_DEV = "ShowAllActiveOrInactive";
        private const string SHOW_ALL_PATCH = "//*[@id=\"ShowActive\"][@value='All']";

        private const string SHOW_ACTIVE_DEV = "ShowOnlyActive";
        private const string SHOW_ACTIVE_PATCH = "//*[@id=\"ShowActive\"][@value='ActiveOnly']";

        private const string SHOW_INACTIVE_DEV = "ShowOnlyInactive";
        private const string SHOW_INACTIVE_PATCH = "//*[@id=\"ShowActive\"][@value='InactiveOnly']";

        private const string SHOW_ALL_AUTO_DEV = "ShowAllAutoOrNot";
        private const string SHOW_ALL_AUTO_PATCH = "//*[@id=\"MenuDisplayOption\"][@value='ShowAll']";

        private const string SHOW_EXCEPT_AUTO_DEV = "ShowAllExceptAuto";
        private const string SHOW_EXCEPT_AUTO_PATCH = "//*[@id=\"MenuDisplayOption\"][@value='AllExceptStay']";

        private const string SHOW_ONLY_AUTO_DEV = "ShowOnlyAuto";
        private const string SHOW_ONLY_AUTO_PATCH = "//*[@id=\"MenuDisplayOption\"][@value='StayOnly']";

        private const string DATE_FROM = "date-picker-start";
        private const string DATE_TO = "date-picker-end";
        private const string SITE_FILTER = "SelectedSites_ms";
        private const string UNCHECK_ALL_SITES = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SEARCH_SITES = "/html/body/div[10]/div/div/label/input";
        private const string TYPE_GUEST_FILTER = "SelectedGuestTypes_ms";
        private const string UNCHECK_ALL_TYPE_GUESTS = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_TYPE_GUESTS = "/html/body/div[11]/div/div/label/input";
        private const string TYPE_MEAL_FILTER = "SelectedMealTypes_ms";
        private const string UNCHECK_ALL_TYPE_MEALS = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SEARCH_TYPE_MEALS = "/html/body/div[12]/div/div/label/input";
        private const string SERVICE_FILTER = "SelectedServices_ms";
        private const string UNCHECK_ALL_SERVICES = "/html/body/div[13]/div/ul/li[2]/a";
        private const string SEARCH_SERVICES = "/html/body/div[13]/div/div/label/input";

        private const string MENU_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[2]";
        private const string MENU_NAME = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[3]";
        private const string MENU_SITE = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[4]";
        private const string MENU_SITE1 = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[4]";
        private const string MENU_MEAl = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[5]";
        private const string MENU_PERIOD = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[6]";
        private const string MENU_PERIOD1 = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[6]";
        private const string INACTIVE = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[1]/img[@alt='Inactive']";

        private const string MASSIVE_DELETE = "MassiveDelete-Menu-Btn";
        private const string SEARCH_MENUS_BTN = "SearchMenusBtn";
        private const string MENU_PRICE = "//*[@id=\"datasheet-list\"]/div[2]/div/div[2]/div/span[2]";

        private const string SEARCH_RESULT_MENU_SITE = "//*[@id=\"tableMenus\"]/tbody/tr[*]/td[2]";
        private const string MASSIVE_DELETE_PAGESIZE = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/nav/select";
        private const string SEARCH_RESULT_MENU_NAME = "//*[@id=\"tableMenus\"]/tbody/tr[*]/td[3]";


        private const string PAGE_TWO = "//*[@id=\"tabContentItemContainer\"]/nav/ul/li[4]/a";
        private const string TABLE_ROWS = "//*[@id=\"tableListMenu\"]/tbody/tr";
        private const string ITEM_NAME = "//*[@id=\"tableMenus\"]/tbody/tr[*]/td[3]";
        private const string MENU_NAME_BTN = "//*[@id=\"tableMenus\"]/thead/tr/th[3]/span/a";
        public const string MASSIVEDELETE = "//*[@id=\"MassiveDelete-Menu-Btn\"]";
        public const string EXPORT_MENUS = "exportBtn";

        
        private const string FILL_DATE_FROM = "dateMenuStart";
        private const string FILL_DATE_TO = "dateMenuEnd";
        private const string LIST_DATE_FROM = "//*[@id=\"tableMenus\"]/tbody/tr[*]/td[5]";
        private const string LIST_DATE_TO = "//*[@id=\"tableMenus\"]/tbody/tr[*]/td[6]";
        private const string MENU_NAME_SEARCH = "/html/body/div[3]/div/div/div[2]/div/form/div/div[1]/div/input";
        private const string BTN_DELETE_MENU = "deleteMenusBtn";
        private const string DATA_CONFIRM_OK = "dataConfirmOK";


        // ______________________________________ Variables ______________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = PLUS_BUTTON)]
        private IWebElement _plusButton;

        [FindsBy(How = How.LinkText, Using = NEW_MENU)]
        private IWebElement _createNewMenu;

        // Tableau menu
        [FindsBy(How = How.XPath, Using = FIRST_MENU_NAME)]
        private IWebElement _firstMenuName;

        [FindsBy(How = How.XPath, Using = FIRST_MENU_SITE)]
        private IWebElement _firstMenuSite;

        [FindsBy(How = How.XPath, Using = MENU_PRICE)]
        private IWebElement _menuPrice;

        [FindsBy(How = How.XPath, Using = MASSIVE_DELETE)]
        private IWebElement _massiveDelete;

        // ______________________________________ Filtres ________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = SHOW_ALL_DEV)]
        private IWebElement _showAllDev;

        [FindsBy(How = How.XPath, Using = SHOW_ALL_PATCH)]
        private IWebElement _showAllPatch;

        [FindsBy(How = How.Id, Using = SHOW_ACTIVE_DEV)]
        private IWebElement _showActiveDev;

        [FindsBy(How = How.XPath, Using = SHOW_ACTIVE_PATCH)]
        private IWebElement _showActivePatch;

        [FindsBy(How = How.Id, Using = SHOW_INACTIVE_DEV)]
        private IWebElement _showInactiveDev;

        [FindsBy(How = How.XPath, Using = SHOW_INACTIVE_PATCH)]
        private IWebElement _showInactivePatch;

        [FindsBy(How = How.Id, Using = SHOW_ALL_AUTO_DEV)]
        private IWebElement _showAllAutoPatch;

        [FindsBy(How = How.XPath, Using = SHOW_ALL_AUTO_PATCH)]
        private IWebElement _showAllAutoDev;

        [FindsBy(How = How.Id, Using = SHOW_EXCEPT_AUTO_DEV)]
        private IWebElement _showAllExceptAutoDev;

        [FindsBy(How = How.XPath, Using = SHOW_EXCEPT_AUTO_PATCH)]
        private IWebElement _showAllExceptAutoPatch;

        [FindsBy(How = How.Id, Using = SHOW_ONLY_AUTO_DEV)]
        private IWebElement _showOnlyAutoDev;

        [FindsBy(How = How.XPath, Using = SHOW_ONLY_AUTO_PATCH)]
        private IWebElement _showOnlyAutoPatch;

        [FindsBy(How = How.Id, Using = DATE_FROM)]
        private IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = DATE_TO)]
        private IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SITES)]
        private IWebElement _uncheckAllSites;

        [FindsBy(How = How.XPath, Using = SEARCH_SITES)]
        private IWebElement _searchSite;

        [FindsBy(How = How.Id, Using = TYPE_GUEST_FILTER)]
        private IWebElement _typeGuestFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_TYPE_GUESTS)]
        private IWebElement _uncheckAllTypeGuests;

        [FindsBy(How = How.XPath, Using = SEARCH_TYPE_GUESTS)]
        private IWebElement _searchTypeGuest;

        [FindsBy(How = How.Id, Using = TYPE_MEAL_FILTER)]
        private IWebElement _typeMealFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_TYPE_MEALS)]
        private IWebElement _uncheckAllTypeMeals;

        [FindsBy(How = How.XPath, Using = SEARCH_TYPE_MEALS)]
        private IWebElement _searchTypeMeal;

        [FindsBy(How = How.Id, Using = SERVICE_FILTER)]
        private IWebElement _serviceFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SERVICES)]
        private IWebElement _uncheckAllServices;

        [FindsBy(How = How.XPath, Using = SEARCH_SERVICES)]
        private IWebElement _searchService;

        [FindsBy(How = How.XPath, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = MENU_NAME_BTN)]
        private IWebElement _menuNameBtn;
        [FindsBy(How = How.Id, Using = FILL_DATE_FROM)]
        private IWebElement _fillDateFrom;
        [FindsBy(How = How.Id, Using = FILL_DATE_TO)]
        private IWebElement _fillDateTo;
        [FindsBy(How = How.Id, Using = MENU_NAME_SEARCH)]
        private IWebElement _menuname;

        public enum FilterType
        {
            SearchMenu,
            SortBy,
            ShowAll,
            ShowActive,
            ShowInactive,
            ShowAllAuto,
            ShowAllExceptAuto,
            ShowOnlyAuto,
            DateFrom,
            DateTo,
            Site,
            TypeGuest,
            TypeMeal,
            Service
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.SearchMenu:
                    _searchFilter = WaitForElementIsVisibleNew(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisibleNew(By.Id(SORTBY_FILTER));
                    _sortBy.Click();
                    var element = WaitForElementIsVisibleNew(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;
                case FilterType.ShowAll:
                    try
                    {
                        _showAllDev = WaitForElementExists(By.Id(SHOW_ALL_DEV));
                        action.MoveToElement(_showAllDev).Perform();
                        _showAllDev.SetValue(ControlType.RadioButton, value);
                    }
                    catch
                    {
                        _showAllPatch = WaitForElementExists(By.XPath(SHOW_ALL_PATCH));
                        action.MoveToElement(_showAllPatch).Perform();
                        _showAllPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.ShowActive:
                    try
                    {
                        _showActiveDev = WaitForElementExists(By.Id(SHOW_ACTIVE_DEV));
                        action.MoveToElement(_showActiveDev).Perform();
                        _showActiveDev.SetValue(ControlType.RadioButton, value);
                    }
                    catch
                    {
                        _showActivePatch = WaitForElementExists(By.XPath(SHOW_ACTIVE_PATCH));
                        action.MoveToElement(_showActivePatch).Perform();
                        _showActivePatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.ShowInactive:
                    try
                    {
                        _showInactiveDev = WaitForElementExists(By.Id(SHOW_INACTIVE_DEV));
                        action.MoveToElement(_showInactiveDev).Perform();
                        _showInactiveDev.SetValue(ControlType.RadioButton, value);
                    }
                    catch
                    {
                        _showInactivePatch = WaitForElementExists(By.XPath(SHOW_INACTIVE_PATCH));
                        action.MoveToElement(_showInactivePatch).Perform();
                        _showInactivePatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.ShowAllAuto:
                    try
                    {
                        _showAllAutoDev = WaitForElementExists(By.Id(SHOW_ALL_AUTO_DEV));
                        action.MoveToElement(_showAllAutoDev).Perform();
                        _showAllAutoDev.SetValue(ControlType.RadioButton, value);
                    }
                    catch
                    {
                        _showAllAutoPatch = WaitForElementExists(By.XPath(SHOW_ALL_AUTO_PATCH));
                        action.MoveToElement(_showAllAutoPatch).Perform();
                        _showAllAutoPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.ShowAllExceptAuto:
                    try
                    {
                        _showAllExceptAutoDev = WaitForElementExists(By.Id(SHOW_EXCEPT_AUTO_DEV));
                        action.MoveToElement(_showAllExceptAutoDev).Perform();
                        _showAllExceptAutoDev.SetValue(ControlType.RadioButton, value);
                    }
                    catch
                    {
                        _showAllExceptAutoPatch = WaitForElementExists(By.XPath(SHOW_EXCEPT_AUTO_PATCH));
                        action.MoveToElement(_showAllExceptAutoPatch).Perform();
                        _showAllExceptAutoPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.ShowOnlyAuto:
                    try
                    {
                        _showOnlyAutoDev = WaitForElementExists(By.Id(SHOW_ONLY_AUTO_DEV));
                        action.MoveToElement(_showOnlyAutoDev).Perform();
                        _showOnlyAutoDev.SetValue(ControlType.RadioButton, value);
                    }
                    catch
                    {
                        _showOnlyAutoPatch = WaitForElementExists(By.XPath(SHOW_ONLY_AUTO_PATCH));
                        action.MoveToElement(_showOnlyAutoPatch).Perform();
                        _showOnlyAutoPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.DateFrom:
                    _dateFrom = WaitForElementIsVisibleNew(By.Id(DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;
                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisibleNew(By.Id(DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;
                case FilterType.Site:
                    _siteFilter = WaitForElementExists(By.Id(SITE_FILTER));
                    action.MoveToElement(_siteFilter).Perform();
                    _siteFilter.Click();

                    _uncheckAllSites = _webDriver.FindElement(By.XPath(UNCHECK_ALL_SITES));
                    _uncheckAllSites.Click();

                    IWebElement selectedSite;
                    _searchSite = WaitForElementIsVisibleNew(By.XPath(SEARCH_SITES));
                    _searchSite.SetValue(ControlType.TextBox, value + " - " + value);

                    selectedSite = _webDriver.FindElement(By.XPath("//span[text()='" + value + " - " + value + "']"));
                    selectedSite.SetValue(ControlType.CheckBox, true);

                    _siteFilter.Click();
                    break;
                case FilterType.TypeGuest:
                    _typeGuestFilter = WaitForElementExists(By.Id(TYPE_GUEST_FILTER));
                    action.MoveToElement(_typeGuestFilter).Perform();
                    _typeGuestFilter.Click();

                    _uncheckAllTypeGuests = WaitForElementIsVisibleNew(By.XPath(UNCHECK_ALL_TYPE_GUESTS));
                    _uncheckAllTypeGuests.Click();

                    _searchTypeGuest = WaitForElementIsVisibleNew(By.XPath(SEARCH_TYPE_GUESTS));
                    _searchTypeGuest.SetValue(ControlType.TextBox, value);

                    var guestToCheck = WaitForElementIsVisibleNew(By.XPath("//span[text()='" + value + "']"));
                    guestToCheck.SetValue(ControlType.CheckBox, true);

                    _typeGuestFilter.Click();
                    break;
                case FilterType.TypeMeal:
                    _typeMealFilter = WaitForElementExists(By.Id(TYPE_MEAL_FILTER));
                    action.MoveToElement(_typeMealFilter).Perform();
                    _typeMealFilter.Click();

                    _uncheckAllTypeMeals = WaitForElementIsVisibleNew(By.XPath(UNCHECK_ALL_TYPE_MEALS));
                    _uncheckAllTypeMeals.Click();

                    _searchTypeMeal = WaitForElementIsVisibleNew(By.XPath(SEARCH_TYPE_MEALS));
                    _searchTypeMeal.SetValue(ControlType.TextBox, value);

                    var mealToCheck = WaitForElementIsVisibleNew(By.XPath("//span[text()='" + value + "']"));
                    mealToCheck.SetValue(ControlType.CheckBox, true);

                    _typeMealFilter.Click();
                    break;
                case FilterType.Service:
                    _serviceFilter = WaitForElementExists(By.Id(SERVICE_FILTER));
                    action.MoveToElement(_serviceFilter).Perform();
                    _serviceFilter.Click();

                    _uncheckAllServices = _webDriver.FindElement(By.XPath(UNCHECK_ALL_SERVICES));
                    _uncheckAllServices.Click();

                    _searchService = WaitForElementIsVisibleNew(By.XPath(SEARCH_SERVICES));
                    _searchService.SetValue(ControlType.TextBox, value);

                    var serviceToCheck = WaitForElementIsVisibleNew(By.XPath("//span[text()='" + value + "']"));
                    serviceToCheck.SetValue(ControlType.CheckBox, true);

                    _serviceFilter.Click();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(FilterType.DateTo, DateUtils.Now);
            }
        }

        public void ClearDateToFilter()
        {
            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateTo.ClearElement();
            _dateTo.SendKeys(Keys.Tab);
        }

        public void ClearDateFromFilter()
        {
            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _dateFrom.ClearElement();
            _dateFrom.SendKeys(Keys.Tab);
        }

        public void UnselectServicesFromMenu()
        {
            Actions action = new Actions(_webDriver);

            _serviceFilter = WaitForElementExists(By.Id(SERVICE_FILTER));
            action.MoveToElement(_serviceFilter).Perform();
            _serviceFilter.Click();

            _uncheckAllServices = _webDriver.FindElement(By.XPath(UNCHECK_ALL_SERVICES));
            _uncheckAllServices.Click();

            WaitPageLoading();
            Thread.Sleep(1500);
        }

        public bool IsSortedById()
        {
            bool valueBool = true;
            double ancienNumber = Double.MinValue;
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

            for (int i = 1; i < tot; i++)
            {
                try
                {
                    IWebElement element = WaitForElementIsVisible(By.XPath(string.Format(MENU_NUMBER, i)));
                    string name = element.Text.Substring(element.Text.IndexOf("°") + 1).Trim();
                    double number = Double.Parse(name);

                    if (i == 1)
                        ancienNumber = number;

                    if (ancienNumber < number)
                    { valueBool = false; }

                    ancienNumber = number;
                }
                catch
                {
                    valueBool = false;
                }
            }
            return valueBool;
        }

        public bool IsSortedByStartDate()
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            var dateFormat = _dateFrom.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? ci = new CultureInfo("fr-FR") : ci = new CultureInfo("en-US");
            int i = 1;
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


            var elements = _webDriver.FindElements(By.XPath(MENU_PERIOD1));

            foreach (var element in elements)
            {
                try
                {

                    string dateText = element.Text.Substring(0, element.Text.IndexOf("-")).Trim();
                    DateTime date = DateTime.Parse(dateText, ci);

                    if (i == 1)
                    {
                        ancientDate = date;
                    }

                    if (DateTime.Compare(ancientDate.Date, date) < 0)
                    { valueBool = false; }

                    ancientDate = date;
                }
                catch
                {
                    valueBool = false;
                }
                i++;
            }

            return valueBool;
        }

        public bool IsSortedByEndDate()
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            var dateFormat = _dateFrom.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? ci = new CultureInfo("fr-FR") : ci = new CultureInfo("en-US");
            int i = 1;
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


            var elements = _webDriver.FindElements(By.XPath(MENU_PERIOD1));

            foreach (var element in elements)
            {
                try
                {

                    string dateText = element.Text.Substring(element.Text.IndexOf("-") + 1).Trim();
                    DateTime date = DateTime.Parse(dateText, ci);

                    if (i == 1)
                    {
                        ancientDate = date;
                    }

                    if (DateTime.Compare(ancientDate.Date, date) < 0)
                    { valueBool = false; }

                    ancientDate = date;
                }
                catch
                {
                    valueBool = false;
                }
                i++;
            }
            return valueBool;
        }

        public bool CheckStatus(bool active)
        {
            bool isActive = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    _webDriver.FindElement(By.XPath(String.Format(INACTIVE, i + 1)));

                    if (active)
                        return false;
                }
                catch
                {
                    isActive = true;
                    if (!active)
                        return true;
                }
            }
            return isActive;
        }

        public Boolean IsDateRespected(DateTime fromDate, DateTime toDate)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            var dateFormat = _dateFrom.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? ci = new CultureInfo("fr-FR") : ci = new CultureInfo("en-US");

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

            var elements = _webDriver.FindElements(By.XPath(MENU_PERIOD1));

            foreach (var element in elements)
            {
                try
                {
                    string endDateText = element.Text.Substring(element.Text.IndexOf("-") + 1).Trim();
                    DateTime menuEndDate = DateTime.Parse(endDateText, ci);

                    string fromDateText = element.Text.Substring(0, element.Text.IndexOf("-") - 1);
                    DateTime menuFromDate = DateTime.Parse(fromDateText, ci);

                    if ((menuEndDate.Date >= fromDate.Date || menuFromDate.Date >= fromDate.Date)
                        || (menuEndDate.Date <= toDate.Date || menuEndDate.Date <= toDate.Date)
                        || (toDate != default && fromDate != default && menuFromDate.Date >= fromDate.Date || menuEndDate.Date <= toDate.Date))
                    {
                        valueBool = true;
                    }
                }
                catch
                {
                    valueBool = false;
                }

            }
            return valueBool;
        }

        public Boolean VerifySite(string value)
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

            var elements = _webDriver.FindElements(By.XPath(MENU_SITE1));

            foreach (var element in elements)
            {
                if (element.Text != value)
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }


        public Boolean VerifyMeal(string value)
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

            var elements = _webDriver.FindElements(By.XPath(MENU_MEAl));

            foreach (var element in elements)
            {
                if (!element.Text.Contains(value) )
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        // ______________________________________ Méthodes ______________________________________________

        public MenusPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // Général
        public override void ShowPlusMenu()
        {
            Actions action = new Actions(_webDriver);
            _plusButton = WaitForElementIsVisibleNew(By.XPath(PLUS_BUTTON));
            action.MoveToElement(_plusButton).Perform();
        }

        public MenusCreateModalPage MenuCreatePage()
        {
            ShowPlusMenu();
            WaitForLoad();
            WaitPageLoading();
            _createNewMenu = WaitForElementIsVisibleNew(By.LinkText(NEW_MENU));
            _createNewMenu.Click();
            WaitForLoad();

            return new MenusCreateModalPage(_webDriver, _testContext);
        }

        // Tableau menus
        public MenusDayViewPage SelectFirstMenu()
        {
            WaitForLoad();
            _firstMenuName = WaitForElementIsVisible(By.XPath(FIRST_MENU_NAME));
            _firstMenuName.Click();
            WaitForLoad();

            return new MenusDayViewPage(_webDriver, _testContext);
        }
        public MenusDayViewPage SelectSecondMenu()
        {
            WaitForLoad();
            var _secondMenuName = WaitForElementIsVisible(By.XPath(SECOND_MENU_NAME));
            _secondMenuName.Click();
            WaitForLoad();

            return new MenusDayViewPage(_webDriver, _testContext);
        }
        public string GetFirstMenuName()
        {
            _firstMenuName = WaitForElementIsVisible(By.XPath(FIRST_MENU_NAME));
            return _firstMenuName.Text.Trim();
        }

        public string GetSecondMenuName()
        {
            var _secondMenuName = WaitForElementIsVisible(By.XPath(SECOND_MENU_NAME));
            return _secondMenuName.Text;
        }

        public string GetFirstMenuSite()
        {
            _firstMenuSite = WaitForElementIsVisible(By.XPath(FIRST_MENU_SITE));
            return _firstMenuSite.Text;
        }
        public MenusDayViewPage ClickFirstMenu()
        {
            WaitForLoad();
            WaitForElementIsVisible(By.XPath(FIRST_MENU_NAME)).Click();
            WaitForLoad();
            return new MenusDayViewPage(_webDriver, _testContext);
        }
        public double GetMenuPrice(string currency, string decimalSeparator)
        {
            // A déplacer dans autre classe
            CultureInfo ci = decimalSeparator.Equals(",") ? ci = new CultureInfo("fr-FR") : ci = new CultureInfo("en-US");

            _menuPrice = WaitForElementIsVisible(By.XPath(MENU_PRICE));
            var analyseText = _menuPrice.Text;

            var format = (NumberFormatInfo)ci.NumberFormat.Clone();
            format.CurrencySymbol = currency;
            var mynumber = Decimal.Parse(analyseText, NumberStyles.Currency, format);

            return Convert.ToDouble(mynumber, ci);
        }
        public void MassiveDeleteMenus(string menuName, string site, string variant)
        {
            ShowExtendedMenu();
            _massiveDelete = WaitForElementIsVisibleNew(By.Id(MASSIVE_DELETE));
            _massiveDelete.Click();
            WaitForLoad();
            var searchPattern = WaitForElementIsVisibleNew(By.XPath("//*[@id=\"formMassiveDeleteMenu\"]/div/div[1]/div/input[@id=\"SearchPattern\"]"));
            searchPattern.Click();
            searchPattern.SendKeys(menuName);
            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", site,false));
            ComboBoxSelectById(new ComboBoxOptions("SelectedVariants_ms", variant, false));
            var searchMenusbtn = WaitForElementIsVisibleNew(By.Id(SEARCH_MENUS_BTN));
            searchMenusbtn.Click();
            WaitForLoad();
            if (isElementVisible(By.XPath("//*[@id=\"tableMenus\"]/tbody/tr")))
            {
                var checkInput = WaitForElementIsVisibleNew(By.XPath("//*[@id=\"tableMenus\"]/tbody/tr/td[1]"));
                checkInput.Click();
                WaitForLoad();
            }
            var deleteBtn = WaitForElementIsVisibleNew(By.Id(BTN_DELETE_MENU));
            deleteBtn.Click();
            WaitForLoad();
            WaitLoading();
            var deleteConfirmBtn = WaitForElementIsVisibleNew(By.Id(DATA_CONFIRM_OK));
            deleteConfirmBtn.Click();
            WaitLoading();
            WaitForLoad();
            var okButton = WaitForElementIsVisibleNew(By.XPath("//*/button[text()='Ok']"));
            okButton.Click();
            WaitForLoad();
        }
        public void MassiveDeleteSiteSearch(string site)
        {
            ShowExtendedMenu();
            _massiveDelete = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE));
            _massiveDelete.Click();
            WaitForLoad();

            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", site,false));

            var searchMenusbtn = WaitForElementIsVisible(By.Id(SEARCH_MENUS_BTN));
            searchMenusbtn.Click();
            WaitForLoad();
        }
        public bool VerifySiteInMassiveDeleteSearch(string value)
        {
            bool valueBool = true;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                if (!isPageSizeEqualsTo100())
                {
                    PageSizeForMassiveDelete("100");
                }
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            var elements = _webDriver.FindElements(By.XPath(SEARCH_RESULT_MENU_SITE));

            foreach (var element in elements)
            {
                if (element.Text != value)
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }
        public new void PageSizeForMassiveDelete(string size)

        {

            string pagesizetionxpath = MASSIVE_DELETE_PAGESIZE;

            if (!isElementExists(By.XPath(pagesizetionxpath)))

            {

                pagesizetionxpath = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/div/nav/select";

            }

            var PageSizeDdl = _webDriver.FindElement(By.XPath(pagesizetionxpath));

            PageSizeDdl.SetValue(ControlType.DropDownList, size);

            WaitForLoad();

        }


        public void MassiveDeleteNameSearch(string menuName)
        {
            ShowExtendedMenu();
            _massiveDelete = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE));
            _massiveDelete.Click();
            WaitForLoad();

            var searchPattern = WaitForElementIsVisible(By.XPath("//*[@id=\"formMassiveDeleteMenu\"]/div/div[1]/div/input[@id=\"SearchPattern\"]"));
            searchPattern.SendKeys(menuName);

            var searchMenusbtn = WaitForElementIsVisible(By.Id(SEARCH_MENUS_BTN));
            searchMenusbtn.Click();
            WaitForLoad();
        }
        public bool VerifyMenuNameInMassiveDeleteSearch(string value)
        {
            bool valueBool = true;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                if (!isPageSizeEqualsTo100())
                {
                    PageSizeForMassiveDelete("100");
                }
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            var elements = _webDriver.FindElements(By.XPath(SEARCH_RESULT_MENU_NAME));

            foreach (var element in elements)
            {
                if (element.Text != value)
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }
        public string GetNameSearched()
        {
            _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
            return _searchFilter.GetAttribute("value");
        }
        public string GetIsActiveOrInactiveChecked()
        {
            var elements = _webDriver.FindElements(By.XPath("//*[@id=\"menuFilterForm\"]/div[4]/input"));
            string selectedShow = string.Empty; 
            foreach (var element in elements)
            {
                if (element.Selected)
                {
                     selectedShow= element.GetAttribute("value");
                }
            }
            return selectedShow;
        }
        public List<string> GetSelectedSitesFilter()
        {
            List<string> selectedFilter = new List<string>();

            _siteFilter = WaitForElementExists(By.Id(SITE_FILTER));
            _siteFilter.Click();
            WaitForLoad();
            var sites = _webDriver.FindElements(By.XPath("/html/body/div[10]/ul/li[*]/label/input"));
            for (var i = 0; i < sites.Count; i++)
            {
                if (sites[i].Selected)
                {
                    var element = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[10]/ul/li[*]/label/input/../span", i + 1)));
                    selectedFilter.Add(element.Text);
                }
            }
            _siteFilter.Click();
            return selectedFilter;
        }
        public List<string> GetSelectedTypeGuestFilter()
        {
            List<string> selectedFilter = new List<string>();

            _typeGuestFilter = WaitForElementExists(By.Id(TYPE_GUEST_FILTER));
            new Actions(_webDriver).MoveToElement(_typeGuestFilter).Perform();
            _typeGuestFilter.Click();
            WaitForLoad();
            var typeGuest = _webDriver.FindElements(By.XPath("/html/body/div[11]/ul/li[*]/label/input"));
            for (var i = 0; i < typeGuest.Count; i++)
            {
                if (typeGuest[i].Selected)
                {
                    var element = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[11]/ul/li[*]/label/input/../span", i + 1)));
                    selectedFilter.Add(element.Text);
                }
            }
            _typeGuestFilter.Click();
            return selectedFilter;
        }
        public List<string> GetSelectedTypeMealFilter()
        {
            List<string> selectedFilter = new List<string>();

            _serviceFilter = WaitForElementExists(By.Id(SERVICE_FILTER));
            new Actions(_webDriver).MoveToElement(_serviceFilter).Perform();
            _serviceFilter.Click();
            WaitForLoad();
            var services = _webDriver.FindElements(By.XPath("/html/body/div[13]/ul/li[*]/label/input"));
            for (var i = 0; i < services.Count; i++)
            {
                if (services[i].Selected)
                {
                    var element = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[13]/ul/li[*]/label/input/../span", i + 1)));
                    selectedFilter.Add(element.Text);
                }
            }
            WaitPageLoading();
            _serviceFilter.Click();
            return selectedFilter;
        }
        public List<string> GetSelectedServiceFilter()
        {
            List<string> selectedFilter = new List<string>();

            _typeMealFilter = WaitForElementExists(By.Id(TYPE_MEAL_FILTER));
            _typeMealFilter.Click();
            WaitForLoad();
            var typesMeal = _webDriver.FindElements(By.XPath("/html/body/div[12]/ul/li[*]/label/input"));
            for (var i = 0; i < typesMeal.Count; i++)
            {
                if (typesMeal[i].Selected)
                {
                    var element = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[12]/ul/li[*]/label/input/../span", i + 1)));
                    selectedFilter.Add(element.Text);
                }
            }
            var serviceFilter = WaitForElementIsVisible(By.Id(SERVICE_FILTER));
            serviceFilter.Click();
            return selectedFilter;
        }
        public int GetNumberOfMenus()
        {
            var tableOfMenu = _webDriver.FindElements(By.XPath(TABLE_ROWS));
            return tableOfMenu.Count;
        }

        public void GoToPageTwo()
        {
            var _pageTwo = WaitForElementExists(By.XPath(PAGE_TWO));
            _pageTwo.Click();
            WaitForLoad();
        }

        public MenuMassiveDeletePage GoToMassiveDelete()
        {
             ShowExtendedMenu();
            WaitForLoad();
             var massiveDeleteElement = WaitForElementExists(By.XPath(MASSIVEDELETE));
            massiveDeleteElement.Click();


             WaitForLoad();

             return new MenuMassiveDeletePage(_webDriver, _testContext);
        }
        public void ExportMenus()
        {
            ShowExtendedMenu();
            var exportMenus = WaitForElementIsVisibleNew(By.Id(EXPORT_MENUS));
            exportMenus.Click();
            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            ClickPrintButton();

            WaitForDownload();
            Close();
        }
        public RecipeMassiveDeletePopup OpenMassiveDeletePopup()
        {
            ShowExtendedMenu();

            var massiveDelete = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE));
            massiveDelete.Click();
            WaitPageLoading();
            return new RecipeMassiveDeletePopup(_webDriver, _testContext);
        }
        public enum FillType
        {
            menuname,
            DateFrom,
            DateTo

        }
        public void Fill(FillType fillType, object value)
        {
            Actions action = new Actions(_webDriver);
            switch (fillType)
            {
                case FillType.menuname:
                    _menuname = WaitForElementIsVisible(By.XPath(MENU_NAME_SEARCH));
                    _menuname.SetValue(ControlType.TextBox, value);
                    _menuname.SendKeys(Keys.Tab);
                    break;

                case FillType.DateFrom:
                    _fillDateFrom = WaitForElementIsVisible(By.Id(FILL_DATE_FROM));
                    _fillDateFrom.SetValue(ControlType.DateTime, (DateTime)value);
                    _fillDateFrom.SendKeys(Keys.Tab);
                    break;

                case FillType.DateTo:
                    _fillDateTo = WaitForElementIsVisible(By.Id(FILL_DATE_TO));
                    _fillDateTo.SetValue(ControlType.DateTime, (DateTime)value);
                    _fillDateTo.SendKeys(Keys.Tab);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FillType), fillType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void ClickSearch()
        {
            var _clickSearc = WaitForElementExists(By.Id(SEARCH_MENUS_BTN));
            _clickSearc.Click();
            WaitForLoad();
        }

        public List<string> GetListDateFrom()
        {
            var ids = _webDriver.FindElements(By.XPath(LIST_DATE_FROM));
            return ids.Select(e => e.Text).ToList();
        }
        public List<string> GetListDateTo()
        {
            var ids = _webDriver.FindElements(By.XPath(LIST_DATE_TO));
            return ids.Select(e => e.Text).ToList();
        }
        public bool IsDateFromGreaterOrEqualToAll(List<string> dateFromlist, DateTime dateFrom)
        {
            foreach (var dateString in dateFromlist)
            {
                DateTime date;
                if (DateTime.TryParse(dateString, out date))
                {
                    if (date >= dateFrom)

                    {
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }
            return true;
        }
        public bool IsDateToGreaterOrEqualToAll(List<string> dateToList, DateTime dateTo)
        {
            foreach (var dateString in dateToList)
            {
                DateTime date;
                if (DateTime.TryParse(dateString, out date))
                {
                    if (date >= dateTo)

                    {
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }
            return true;
        }
        public MenusMassiveDeletePopup MenuOpenMassiveDeletePopup()
        {
            ShowExtendedMenu();
            var massiveDeleteElement = WaitForElementExists(By.XPath(MASSIVEDELETE));
            massiveDeleteElement.Click();
            LoadingPage();


            WaitForLoad();

            return new MenusMassiveDeletePopup(_webDriver, _testContext);
        }
        public bool IsExcelFileCorrect(string filePath)
        {
            Regex r = new Regex("[export?flights\\s\\d.-]", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);
            return m.Success;
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            StringBuilder sb = new StringBuilder();

            foreach (var file in taskFiles)
            {
                sb.Append(file.Name + " ");
                //  Test REGEX
                if (IsExcelFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count == 0)
            {
                return null;
            }
            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }
        public List<string> GetFilteredMenuList()
        {
            List<string> filteredMenus = new List<string>();

            // Locate the menu elements on the page after filtering
            var menuElements = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody"));

            // Loop through the elements and extract menu names
            foreach (var element in menuElements)
            {
                filteredMenus.Add(element.Text.Trim());
            }

            return filteredMenus;
        }

        public bool IsConceptUnique()
        {
            WaitPageLoading();
            var conceptField = _webDriver.FindElement(By.XPath("//*[@id=\"dropdown-Concept\"]"));

            // Obtenir la valeur sélectionnée du menu déroulant Concept
            var selectedConcept = new SelectElement(conceptField).SelectedOption.Text;

            // Vérifier les doublons 
            var allConcepts = _webDriver.FindElements(By.XPath("//*[@id=\"dropdown-Concept\"]/option"));
        

            // Si le concept apparaît plus d'une fois, retourner false (dupliqué), sinon true (unique)
            return allConcepts.Select(option => option.Text).Distinct().Count() == allConcepts.Count(); 
        }

    }
}
