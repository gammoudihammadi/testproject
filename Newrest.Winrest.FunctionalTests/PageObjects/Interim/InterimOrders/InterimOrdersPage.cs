using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Commons.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimReceptions;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders
{
    public class InterimOrdersPage : PageBase
    {
        //__________________________________Constantes_______________________________________

        // General
        private const string NEW_INTERIM_ORDER_BUTTON = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[2]/div/a[text()='New interim order']";
        private const string FIRST_INTERIM = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[2]";
        private const string SECOND_INTERIM = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[2]/td[2]";
        private const string NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]";
        private const string SUPPLIER_VALIDATOR_MESSAGE = "//*[@id=\"form-create-interim-order\"]/div/div[3]/div[1]/div/div/span]";
        private const string LOCATION_VALIDATOR_MESSAGE = "//*[@id=\"form-create-interim-order\"]/div/div[3]/div[2]/div/div/span]";
        private const string LIST_SUPPLIER_FILTER = "/html/body/div[10]/ul/li[*]/label/input";
        private const string RESET_FILTER = "//*[@id=\"formSearchInterimOrders\"]/div[1]/a";
        private const string SHOW_ALL_INTERIM_ORDERS = "show-only-all";
        private const string SHOW_NOT_VALIDATED_ONLY = "show-only-not-validated";
        private const string SHOW_VALIDATED_ONLY = "show-only-validated";
        private const string FIRST_INTERIM_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[2]";
        private const string FILTER_SEARCH_NUMBER = "SearchNumber";
        private const string FILTER_INTERIM_RECEPTION_STATUS_ALL = "//*[@id=\"status-all\"]";
        private const string FILTER_INTERIM_RECEPTION_STATUS_OPENED = "//*[@id=\"status-opened\"]";
        private const string FILTER_INTERIM_RECEPTION_STATUS_CLOSED = "//*[@id=\"status-closed\"]";
        private const string FILTER_INTERIM_RECEPTION_STATUS_CANCELLED = "//*[@id=\"status-cancelled\"]";
        private const string LIST_DELIVERY_DATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[5]";
        private const string FILTER_DATE_FROM = "date-picker-start";
        private const string FILTER_DATE_TO = "date-picker-end";
        private const string SHOW_SELECTED = "ShowActive";
        private const string SUPPLIERS = "SelectedSuppliers_ms";
        private const string SUPPLIERS_SELECTED = "/html/body/div[10]/ul/li/label/input";
        private const string SHOW_ACTIVE_SELECTED = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[8]/input";
        private const string Sites_FILTER = "SelectedSites_ms";
        private const string SHOW_STATUS_SELECTED = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[10]/div/input";
        private const string SITES_SELECTED = "/html/body/div[11]/ul/li/label/input";
        private const string FILTER_SUPPLIER = "SelectedSuppliers_ms";
        private const string SITE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[3]";
        private const string SEARCH_SITE = "/html/body/div[11]/div/div/label/input";
        private const string SITES_SELECT_ALL = "/html/body/div[11]/div/ul/li[1]/a/span[2]";
        private const string SITE_UNSELECT_ALL = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string SUPPLIER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[4]";
        private const string LIST_SITE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[3]";
        private const string LIST_SITE_FILTER = "/html/body/div[11]/ul/li[*]/label/input";
        private const string BY_FROM_TO_DATE = "SearchModeDateValue";
        private const string BY_VALIDATION_DATE = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[6]/input[2]";
        private const string DETAIL_BUTTON = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[5]";
        private const string QTY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[8]";
        private const string ITEM_INTERIM_ORDER = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]";

        //ShowAll ShowOnlyInactive ShowOnlyActive
        private const string SHOW_ALL = "//*[@id=\"ShowActive\"][@value='All']";
        private const string SHOW_ONLY_ACTIVE = "//*[@id=\"ShowActive\"][@value='ActiveOnly']";
        private const string SHOW_ONLY_INACTIVE = "//*[@id=\"ShowActive\"][@value='InactiveOnly']";
        private const string FIRST_LINE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]";
        private const string PAGINATION_NEXT_FIRST_PAGE = "//*[@id=\"list-item-with-action\"]/nav/ul/li[4]/a";
        private const string PAGINATION_NEXT = "//*[@id=\"list-item-with-action\"]/nav/ul/li[3]/a";
        private const string KEYWORDS = "ItemIndexVMSelectedKeywords_ms";
        private const string UNSELECT_ALL_KEYWORDS = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string SELECT_ALL_KEYWORDS = "/html/body/div[11]/div/ul/li[1]/a/span[2]";
        private const string KEYWORD = "tbSearchByKeyword";
        private const string RESULTS_ROWS = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]";
        private const string DETAIL_BUTTONS = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[2]/td[5]";
        private const string TOTALPEICE = "total-price-span";
        private const string NAME_REF = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[3]/span";
        private const string NAME_REF_PATCH = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[3]/span";
        private const string LISTE_NUMBERS_INTERORDE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[2]";
        private const string KEYWORDS_COMBO = "ItemIndexVMSelectedKeywords_ms";
        private const string FIRST_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[2]";
        private const string NUMBER_ORDERS = "//*[@id=\"list-item-with-action\"]/table";
        private const string ORDERS_ITEMS = "//*[@id=\"list-item-with-action\"]/table/tbody";

        private const string NUMBER_LINE = "/html/body/div[2]/div/div[2]/div[1]/h1/span";
        public InterimOrdersPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        //__________________________________Variables_______________________________________

        // Général

        [FindsBy(How = How.Id, Using = NUMBER)]
        private IWebElement _validateResults;

        [FindsBy(How = How.XPath, Using = NEW_INTERIM_ORDER_BUTTON)]
        private IWebElement _createNewInterim;

        // Tableau
        [FindsBy(How = How.XPath, Using = FIRST_INTERIM)]
        private IWebElement _firstInterim;

        [FindsBy(How = How.Id, Using = SHOW_VALIDATED_ONLY)]
        private IWebElement _showvalidatedonly;

        [FindsBy(How = How.Id, Using = SHOW_NOT_VALIDATED_ONLY)]
        private IWebElement _shownotvalidatedonly;

        [FindsBy(How = How.Id, Using = SHOW_ALL_INTERIM_ORDERS)]
        private IWebElement _showallinterimorder;

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_SEARCH_NUMBER)]
        private IWebElement _searchByNumber;

        [FindsBy(How = How.XPath, Using = FILTER_INTERIM_RECEPTION_STATUS_ALL)]
        private IWebElement _interimReceptionsAll;

        [FindsBy(How = How.XPath, Using = FILTER_INTERIM_RECEPTION_STATUS_OPENED)]
        private IWebElement _interimReceptionsOpened;

        [FindsBy(How = How.XPath, Using = FILTER_INTERIM_RECEPTION_STATUS_CLOSED)]
        private IWebElement _interimReceptionsClosed;

        [FindsBy(How = How.XPath, Using = FILTER_INTERIM_RECEPTION_STATUS_CANCELLED)]
        private IWebElement _interimReceptionsCancelled;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;
        [FindsBy(How = How.Id, Using = FILTER_SUPPLIER)]
        private IWebElement _supplierFilter;
        [FindsBy(How = How.Id, Using = Sites_FILTER)]
        private IWebElement _SitesFilter;
        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSites;
        [FindsBy(How = How.XPath, Using = SITES_SELECT_ALL)]
        private IWebElement _siteSelectAll;
        [FindsBy(How = How.XPath, Using = SITE_UNSELECT_ALL)]
        private IWebElement _siteUnselectAll;

        [FindsBy(How = How.Id, Using = BY_FROM_TO_DATE)]
        private IWebElement _byFromToDate;

        [FindsBy(How = How.XPath, Using = BY_VALIDATION_DATE)]
        private IWebElement _byValidationDate;

        [FindsBy(How = How.XPath, Using = FIRST_INTERIM_NUMBER)]
        private IWebElement _firstInterimNumber;

        [FindsBy(How = How.XPath, Using = QTY)]
        private IWebElement _qty;

        [FindsBy(How = How.XPath, Using = ITEM_INTERIM_ORDER)]
        private IWebElement _itemInterimOrder;


        [FindsBy(How = How.XPath, Using = DETAIL_BUTTON)]
        private IWebElement _detail_Button;
        [FindsBy(How = How.XPath, Using = SHOW_ALL)]
        private IWebElement _showAll;

        [FindsBy(How = How.XPath, Using = SHOW_ONLY_ACTIVE)]
        private IWebElement _showOnlyActive;

        [FindsBy(How = How.XPath, Using = SHOW_ONLY_INACTIVE)]
        private IWebElement _showOnlyInactive;

        [FindsBy(How = How.XPath, Using = PAGINATION_NEXT)]
        private IWebElement _paginationNext;

        [FindsBy(How = How.Id, Using = KEYWORDS)]
        private IWebElement _keywords;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_KEYWORDS)]
        private IWebElement _unselectAllKeywords;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_KEYWORDS)]
        private IWebElement _selectAllKeywords;

        [FindsBy(How = How.Id, Using = KEYWORD)]
        private IWebElement _keyword;

        [FindsBy(How = How.XPath, Using = RESULTS_ROWS)]
        private IWebElement _rowsResults;

        [FindsBy(How = How.XPath, Using = NAME_REF)]
        private IWebElement _name_ref;

        [FindsBy(How = How.XPath, Using = FIRST_NUMBER)]
        private IWebElement _first_number;
        public List<string> GetInterimOrdersList()
        {
            var rows = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table"));
            var ordersList = new List<string>();

            foreach (var row in rows)
            {
                var rowText = row.Text;

                var columns = rowText.Split('\n');

                foreach (var column in columns)
                {
                    ordersList.Add(column.Trim());
                }
            }
            if (ordersList.Count > 0)
            {
                ordersList.RemoveAt(0);
            }
            return ordersList;
        }

        public InterimOrdersCreateModalPage CreateNewInterimOrder()
        {
            ShowPlusMenu();

            _createNewInterim = WaitForElementIsVisible(By.XPath(NEW_INTERIM_ORDER_BUTTON), nameof(NEW_INTERIM_ORDER_BUTTON));
            _createNewInterim.Click();
            WaitForLoad();

            return new InterimOrdersCreateModalPage(_webDriver, _testContext);
        }

        public string GetFirstID()
        {
            WaitPageLoading(); 
            if (isElementVisible(By.XPath(FIRST_INTERIM)))
            {
                _firstInterim = WaitForElementExists(By.XPath(FIRST_INTERIM));
                return _firstInterim.Text;
            }
            else
            {
                return null;
            }

        }
        public string GetSecondID()
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(SECOND_INTERIM)))
            {
                _firstInterim = WaitForElementExists(By.XPath(SECOND_INTERIM));
                return _firstInterim.Text;
            }
            else
            {
                return null;
            }

        }

        public bool CheckValidator()
        {
            return isElementExists(By.XPath(LOCATION_VALIDATOR_MESSAGE)) && isElementExists(By.XPath(SUPPLIER_VALIDATOR_MESSAGE));
        }
        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
        }
        public enum FilterType
        {
            SearchByNumber,
            Suppliers,
            ByFromTodate,
            ByValidationDate,
            ShowValidatedOnly,
            ShowNotValidatedOnly,
            ShowAllOrders,
            ShowAllInterimReception,
            ShowOpenedInterimReception,
            ShowClosedInterimReception,
            ShowCancelledInterimReception,
            DateFrom,
            DateTo,
            Sites,
            ByFromToDate,
            ShowAll,
            ShowOnlyInactive,
            ShowOnlyActive,
            Site,
            Keywords,
            Keyword

        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.ShowValidatedOnly:
                    _showvalidatedonly = WaitForElementExists(By.Id(SHOW_VALIDATED_ONLY));
                    _showvalidatedonly.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
                case FilterType.ByFromToDate:
                    _byFromToDate = WaitForElementExists(By.Id(BY_FROM_TO_DATE));
                    _byFromToDate.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;

                case FilterType.ByValidationDate:
                    _byValidationDate = WaitForElementExists(By.XPath(BY_VALIDATION_DATE));
                    _byValidationDate.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;

                case FilterType.ShowNotValidatedOnly:
                    _shownotvalidatedonly = WaitForElementExists(By.Id(SHOW_NOT_VALIDATED_ONLY));
                    _shownotvalidatedonly.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;

                case FilterType.ShowAllOrders:
                    _showallinterimorder = WaitForElementExists(By.Id(SHOW_ALL_INTERIM_ORDERS));
                    _showallinterimorder.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
                case FilterType.SearchByNumber:
                    _searchByNumber = WaitForElementIsVisible(By.Id(FILTER_SEARCH_NUMBER));
                    _searchByNumber.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;

                case FilterType.ShowAllInterimReception:
                    _interimReceptionsAll = WaitForElementExists(By.XPath(FILTER_INTERIM_RECEPTION_STATUS_ALL));
                    action.MoveToElement(_interimReceptionsAll).Perform();
                    _interimReceptionsAll.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowOpenedInterimReception:
                    _interimReceptionsOpened = WaitForElementExists(By.XPath(FILTER_INTERIM_RECEPTION_STATUS_OPENED));
                    action.MoveToElement(_interimReceptionsOpened).Perform();
                    _interimReceptionsOpened.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowClosedInterimReception:
                    _interimReceptionsClosed = WaitForElementExists(By.XPath(FILTER_INTERIM_RECEPTION_STATUS_CLOSED));
                    action.MoveToElement(_interimReceptionsClosed).Perform();
                    _interimReceptionsClosed.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowCancelledInterimReception:
                    _interimReceptionsCancelled = WaitForElementExists(By.XPath(FILTER_INTERIM_RECEPTION_STATUS_CANCELLED));
                    action.MoveToElement(_interimReceptionsCancelled).Perform();
                    _interimReceptionsCancelled.SetValue(ControlType.RadioButton, value);
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
                case FilterType.Site:

                    ComboBoxSelectById(new ComboBoxOptions(Sites_FILTER, (string)value));
                    break;
                case FilterType.Suppliers:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SUPPLIER, (string)value));
                    break;
                case FilterType.ShowAll:
                    _showAll = WaitForElementExists(By.XPath(SHOW_ALL));
                    _showAll.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;

                case FilterType.ShowOnlyActive:
                    _showOnlyActive = WaitForElementExists(By.XPath(SHOW_ONLY_ACTIVE));
                    _showOnlyActive.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;

                case FilterType.ShowOnlyInactive:
                    _showOnlyInactive = WaitForElementExists(By.XPath(SHOW_ONLY_INACTIVE));
                    _showOnlyInactive.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;
                case FilterType.Sites:
                    ScrollUntilElementIsInView(By.Id(Sites_FILTER));
                    if (isElementVisible(By.Id(Sites_FILTER)))
                    {
                        _SitesFilter = WaitForElementIsVisible(By.Id(Sites_FILTER));
                        _SitesFilter.Click();

                        if (value.Equals("ALL SITES"))
                        {
                            _searchSites = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                            _searchSites.SetValue(ControlType.TextBox, "");

                            _siteSelectAll = WaitForElementIsVisible(By.XPath(SITES_SELECT_ALL));
                            _siteSelectAll.Click();
                        }
                        else
                        {
                            _siteUnselectAll = WaitForElementIsVisible(By.XPath(SITE_UNSELECT_ALL));
                            _siteUnselectAll.Click();

                            _searchSites = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                            _searchSites.SetValue(ControlType.TextBox, value);
                            if (isElementVisible(By.XPath("//span[text()='" + value + " " + "-" + " " + value + "']")))
                            {
                                var valueToCheckCustomersType = WaitForElementIsVisible(By.XPath("//span[text()='" + value + " " + "-" + " " + value + "']"));
                                valueToCheckCustomersType.SetValue(ControlType.CheckBox, true);
                            }
                        }
                        _SitesFilter.Click();
                    }
                    break;
                case FilterType.Keyword:
                    _keyword = WaitForElementIsVisible(By.Id(KEYWORD));
                    _keyword.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;

                case FilterType.Keywords:
                    ComboBoxSelectById(new ComboBoxOptions(KEYWORDS, (string)value));
                    WaitPageLoading();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(filterType), filterType, "Invalid filter type provided.");
            }

            WaitPageLoading();
            WaitForLoad();
        }
        public string GetShowFilterSelected()
        {

            var itemSelected = _webDriver.FindElements(By.Id(SHOW_SELECTED));
            var showName = string.Empty;
            foreach (var element in itemSelected)
            {
                if (element.Selected)
                {
                    switch (element.GetAttribute("value"))
                    {
                        case "ActiveOnly":
                            showName = "Show active only";
                            break;
                        case "InactiveOnly":
                            showName = "Show inactive only";
                            break;
                        case "All":
                            showName = "Show all";
                            break;
                        default:
                            showName = "";
                            break;
                    }
                }
            }

            return showName;
        }

        public List<DateTime> GetDeliveryDate(string writingType, string dateFormatPicker)
        {
            List<DateTime> dates = new List<DateTime>();

            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = dateFormatPicker.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var deliveryDates = _webDriver.FindElements(By.XPath(string.Format(LIST_DELIVERY_DATE, writingType)));

            foreach (var elm in deliveryDates)
            {
                dates.Add(DateTime.Parse(elm.Text.Trim(), ci).Date);
            }

            return new List<DateTime>(dates);
        }
        public string GetStatusFilterSelected()
        {
            var itemSelected = _webDriver.FindElements(By.XPath(SHOW_STATUS_SELECTED));
            var showName = string.Empty;
            foreach (var element in itemSelected)
            {
                if (element.Selected)
                {
                    switch (element.GetAttribute("value"))
                    {
                        case "All":
                            showName = "All";
                            break;
                        case "Opened":
                            showName = "Opened";
                            break;
                        case "Closed":
                            showName = "Closed";
                            break;
                        case "Cancelled":
                            showName = "Cancelled";
                            break;
                        default:
                            showName = "";
                            break;
                    }
                }
            }
            return showName;
        }

        public InterimOrdersItem ClickOnFirstService()
        {
            _firstInterimNumber = WaitForElementIsVisible(By.XPath(FIRST_INTERIM_NUMBER));
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].scrollIntoView(true);", _firstInterimNumber);
            WaitForLoad();
            _firstInterimNumber.Click();
            WaitPageLoading();
            WaitForLoad();
            return new InterimOrdersItem(_webDriver, _testContext);
        }
      
        public InterimOrdersItem ClickFirstLine()
        {
            WaitPageLoading(); 
            _detail_Button = WaitForElementIsVisible(By.XPath(DETAIL_BUTTON), nameof(DETAIL_BUTTON));
            _detail_Button.Click();
            WaitForLoad();
            return new InterimOrdersItem(_webDriver, _testContext);

        }
        public InterimOrdersItem ClickSecondtLine()
        {
            _detail_Button = WaitForElementIsVisible(By.XPath(DETAIL_BUTTONS), nameof(DETAIL_BUTTONS));
            _detail_Button.Click();
            WaitForLoad();
            return new InterimOrdersItem(_webDriver, _testContext);

        }
        public bool IsSortedBySuppliers(string supplier)
        {
            var listeSuppliers = _webDriver.FindElements(By.XPath(SUPPLIER));

            if (listeSuppliers.Count == 0)
                return false;


            foreach (var elm in listeSuppliers)
            {
                try
                {
                    if (elm.Text.CompareTo(supplier) < 0)
                        return false;
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }

        public void ScrollUntilSitesFilterIsVisible()
        {
            ScrollUntilElementIsInView(By.Id(Sites_FILTER));
        }
        public bool IsSortedBySites(string site)
        {
            var listSite = _webDriver.FindElements(By.XPath(LIST_SITE));
            if (listSite.Count == 0)
                return false;
            foreach (var elm in listSite)
            {
                try
                {
                    if (elm.Text.CompareTo(site) < 0)
                        return false;

                }
                catch
                {
                    return false;
                }
            }
            return true;
        }
        public InterimOrdersItem GoToInterim_InterimOrdersItem()
        {
            WaitPageLoading();
            ClickFirstInterim();
            WaitForLoad();
            return new InterimOrdersItem(_webDriver, _testContext);
        }
        public void ClickFirstInterim()
        {
            _firstInterim = WaitForElementIsVisible(By.XPath(FIRST_INTERIM));
            _firstInterim.Click();
            WaitForLoad();
        }
        public object GetFilterValue(FilterType filterType)
        {
            switch (filterType)
            {
                case FilterType.ShowAll:
                    var _showAll = WaitForElementExists(By.XPath(SHOW_ALL));
                    return _showAll.Selected;

                case FilterType.ShowOnlyActive:
                    var _showOnlyActive = WaitForElementIsVisible(By.XPath(SHOW_ONLY_ACTIVE));
                    return _showOnlyActive.Selected;

                case FilterType.ShowOnlyInactive:
                    var _showOnlyInactive = WaitForElementIsVisible(By.XPath(SHOW_ONLY_INACTIVE));
                    return _showOnlyInactive.Selected;

                //case FilterType.ShowAllReceptions:
                //    var _showAllReceptions = WaitForElementIsVisible(By.XPath(SHOW_ALL_RECEPTIONS));
                //    return _showAllReceptions.Selected;

                case FilterType.ShowValidatedOnly:
                    var _showValidatedReceptions = WaitForElementIsVisible(By.XPath(SHOW_VALIDATED_ONLY));
                    return _showValidatedReceptions.Selected;

                case FilterType.ShowNotValidatedOnly:
                    var _showNotValidatedReceptions = WaitForElementIsVisible(By.XPath(SHOW_NOT_VALIDATED_ONLY));
                    return _showNotValidatedReceptions.Selected;

                case FilterType.ShowAllInterimReception:
                    var _interimReceptionStatusAll = WaitForElementExists(By.XPath(FILTER_INTERIM_RECEPTION_STATUS_ALL));
                    return _interimReceptionStatusAll.Selected;

                case FilterType.ShowOpenedInterimReception:
                    var _interimReceptionStatusOpened = WaitForElementExists(By.XPath(FILTER_INTERIM_RECEPTION_STATUS_OPENED));
                    return _interimReceptionStatusOpened.Selected;

                case FilterType.ShowClosedInterimReception:
                    var _interimReceptionStatusClosed = WaitForElementExists(By.XPath(FILTER_INTERIM_RECEPTION_STATUS_CLOSED));
                    return _interimReceptionStatusClosed.Selected;
            }
            return null;
        }
        public int GetNumberSelectedSiteFilter()
        {
            var listSitesSelectedFilters = _webDriver.FindElements(By.XPath(LIST_SITE_FILTER));
            var numberSitesSelectedSite = listSitesSelectedFilters
               .Where(p => p.Selected == true).Count();

            return numberSitesSelectedSite;
        }
        public int GetNumberSelectedSupplierFilter()
        {
            var listSelectedSupplierFilters = _webDriver.FindElements(By.XPath(LIST_SUPPLIER_FILTER));
            var numberSelectedSupplier = listSelectedSupplierFilters.Where(p => p.Selected).Count();

            return numberSelectedSupplier;
        }
        public void ScrollUntilResetFilterIsVisible()
        {
            WaitForLoad();
            ScrollUntilElementIsInView(By.XPath(RESET_FILTER));
            WaitForLoad();
        }

        public List<string> GetNumberList()
        {
            var ListNumber = new List<string>();

            var List = _webDriver.FindElements(By.XPath(LISTE_NUMBERS_INTERORDE));

            foreach (var row in List)
            {
                ListNumber.Add(row.Text);
            }

            return ListNumber;
        }

        public void PaginationNextFirstPage()
        {
            if (isElementVisible(By.XPath(PAGINATION_NEXT_FIRST_PAGE))) {
                _paginationNext = WaitForElementIsVisible(By.XPath(PAGINATION_NEXT_FIRST_PAGE));
                _paginationNext.Click();
                WaitForLoad();
            }


        }
        public void PaginationNext()
        {
            if (isElementVisible(By.XPath(PAGINATION_NEXT)))
            {
                _paginationNext = WaitForElementExists(By.XPath(PAGINATION_NEXT));
                _paginationNext.Click();
                WaitForLoad();
            }


        }
        public void Fill(string keywords)
        {
            // combo box keywords (All)
            _keywords = WaitForElementIsVisible(By.Id(KEYWORDS));
            if (keywords == null)
            {
                _keywords.Click();

                _selectAllKeywords = WaitForElementIsVisible(By.XPath(SELECT_ALL_KEYWORDS));
                _selectAllKeywords.Click();
                // on referme
                _keywords.Click();
            }
            else
            {
                ComboBoxSelectById(new ComboBoxOptions(KEYWORDS_COMBO, keywords, false));
                _keywords.Click();
            }

            WaitForLoad();
        }
        public int GetTotalResultRowsPage()
        {
            var rowsResults = _webDriver.FindElements(By.XPath(RESULTS_ROWS));
            WaitForLoad();
            return rowsResults.Count;
        }

        public string GetFirstInterimOrdersName()
        {
            var firstItem = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]"));
            firstItem.Click();
            if (IsDev())
            {
                _name_ref = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[3]/span"));
                return _name_ref.Text;
            }
            else
            {
                _name_ref = _webDriver.FindElement(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[3]/span")); 
                return _name_ref.Text;
            }
        }
        public string GetFirstInterimOrdersNumber()
        {
           
            _first_number = WaitForElementIsVisible(By.XPath(FIRST_NUMBER));
            return _first_number.Text;
        }

        public int GetNumberOfOrders()
        {
            WaitPageLoading();
            var orderRows = _webDriver.FindElements(By.XPath(NUMBER_ORDERS)); 
            return orderRows.Count;
        }

        public InterimOrdersItem ClickOrderAtIndex(int index)
        {
            WaitPageLoading();
            // Get the list of order rows
            var orderRows = _webDriver.FindElements(By.XPath(ORDERS_ITEMS)); 

            // Ensure the index is within bounds
            if (index >= 0 && index < orderRows.Count)
            {
                // Click on the order at the given index
                orderRows[index].Click();

                return new InterimOrdersItem(_webDriver, _testContext);
            }

            throw new IndexOutOfRangeException("The specified index is out of range.");
        }
        public List<string> GetTheFirstTwoValidTotalVAT()
        {
            WaitPageLoading();
        
            var pagination = WaitForElementIsVisible(By.XPath("//*[@id=\"page-size-selector\"]"));
            pagination.Click();
            var paginationNum = WaitForElementIsVisible(By.XPath("//*[@id=\"page-size-selector\"]/option[2]"));
            paginationNum.Click();
            WaitForLoad(); 
            WaitPageLoading();
            List<string> TotalList = GetInterimOrdersList().AsEnumerable().Reverse().ToList();
            List<int> nonZeroIndexes = new List<int>();
            List<string> ids = new List<string>();

            for (int i = 0; i < TotalList.Count; i++)
            {
                
                string[] parts = TotalList[i].Split(' ');
                string amountStr = parts.Last();  
                decimal amount = decimal.Parse(amountStr.Replace("€", "").Trim());

                if (amount > 0)
                {
                    nonZeroIndexes.Add(i);
                    string id = parts[0];  
                    ids.Add(id);
                    if (nonZeroIndexes.Count == 2)
                    {
                        break;
                    }
                }
            }
            return ids;
        }
        public int GetTotalOfOrders()
        {
            WaitPageLoading();
            var orderRowsTotal = WaitForElementIsVisible(By.XPath(NUMBER_LINE));
            return  Int32.Parse(orderRowsTotal.Text);
        }
       
    
        public void ScrollToFilterSearchNumber()
        {
            _webDriver.Manage().Window.Maximize();

            Actions actions = new Actions(_webDriver);
            IWebElement scrollbar = WaitForElementExists((By.ClassName("ps__scrollbar-y")));
            actions.MoveToElement(scrollbar).Perform();

            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)_webDriver;
            jsExecutor.ExecuteScript("arguments[0].style.top = '0px';", scrollbar);
            WaitPageLoading();

        }
    }
}