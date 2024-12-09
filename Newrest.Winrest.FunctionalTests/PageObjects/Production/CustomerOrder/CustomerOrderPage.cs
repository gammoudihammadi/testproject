using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Threading;
using UglyToad.PdfPig.Fonts.TrueType.Names;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder
{
    public class CustomerOrderPage : PageBase
    {

        public CustomerOrderPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_____________________________________

        // Général
        private const string COUNTER = "//*[@id=\"tabContentItemContainer\"]/div/h1/span";
        private const string PLUS_BTN = "//*[@id=\"tabContentItemContainer\"]/div/div/div[2]";
        private const string EXTENDED_BUTTON = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]";
        private const string PRINT_CUSTOMER_ORDER = "btn-print";
        private const string EXPORT = "btn-export-excel";
        private const string SEND_MAIL_BTN = "btn-send-all-by-email";
        private const string SEND = "btn-init-async-send-mail";

        // Tableau customer order
        private const string NEW_CUSTOMER_ORDER = "production-order-createbtn";
        private const string NEW_CUSTOMER_ORDER_ISPATCH = "production-order-createbtn";
        private const string FIRST_CUSTOMER_ORDER_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[contains(text(),'Id')]";
        private const string DELETE_BTN = "//*/span[contains(@class,'trash')]/..";
        private const string DELETE_CONFIRM = "//button[text()='Delete']";

        private const string FIRST_ORDER_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[5]";
        private const string FIRST_EXPEDITION_DATE = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[12]";

        private const string FIRST_CUSTOMER = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[7]";
        private const string FIRST_DELIVERY = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[8]";
        private const string EYE_ICON_GREEN = "/html/body/div[2]/div/div[2]/div/table/tbody/tr/td[3]/div/span";
        private const string EYE_ICON_GRAY = "/html/body/div[2]/div/div[2]/div/table/tbody/tr[1]/td[3]/div/span";

        // Filtres
        private const string RESET_FILTER = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string SEARCH_FILTER = "SearchPattern";
        private const string FLIGHT_NUMBER_FILTER = "searchPatternByFlightNumber";
        private const string SORTBY_FILTER = "cbSortBy";
        private const string FROM_FILTER = "start-date-picker";
        private const string TO_FILTER = "end-date-picker";
        private const string SITE_FILTER = "collapseSelectedSitesFilter";
        private const string UNSELECT_SITE = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SEARCH_SITE = "/html/body/div[10]/div/div/label/input";
        private const string SHOW_ALL_ORDER = "//*[@id=\"ShowOnlyValue\"][@value='All']";
        private const string SHOW_INVOICED_ONLY = "//*[@id=\"ShowOnlyValue\"][@value='InvoicedOnly']";
        private const string SHOW_NOT_INVOICED_ONLY = "//*[@id=\"ShowOnlyValue\"][@value='NotInvoicedOnly']";
        private const string SHOW_STATUS = "//*[@id=\"item-filter-form\"]/div[8]/a";
        private const string STATUS_ALL = "//*[@id=\"Status\"][@value='None']";
        private const string STATUS_OPENED = "//*[@id=\"Status\"][@value='Opened']";
        private const string STATUS_CLOSED = "//*[@id=\"Status\"][@value='Closed']";
        private const string SHOW_INVOICE_STEP = "//*[@id=\"item-filter-form\"]/div[9]/a";
        private const string INVOICE_ALL = "all";
        private const string INVOICE_IN_PROGRESS = "inProgress";
        private const string INVOICE_VALIDATED = "validated";
        private const string INVOICE_INVOICED = "invoiced";
        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string UNSELECT_CUSTOMERS = "/html/body/div[11]/div/ul/li[2]/a";
        private const string CUSTOMER_SEARCH = "/html/body/div[11]/div/div/label/input";

        private const string VALIDATED = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[1]/img[@alt=\"Valid\"]";
        private const string INVOICE = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[2]/span";
        private const string CUSTOMER_ORDER_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[contains(text(),'Id')]";

        private const string CUSTOMER_NAME = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[*]/b";
        private const string ORDER_DATE = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[9]";
        private const string SENT_BY_MAIL = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[*]/span[contains(@class,'envelope')]";

        private const string VERIFIED_Filter = "//*[@id=\"item-filter-form\"]/div[16]/a[@class=\"filterCollapseButton collapsed\"]";
        private const string VERIFIED_ALL = "//*[@id=\"VerifiedStatus\" and @value=\"All\"]";
        private const string VERIFIED_ONLY = "//*[@id=\"VerifiedStatus\" and @value=\"OnlyIsVerified\"]";
        private const string VERIFIED_NOTONLY = "//*[@id=\"VerifiedStatus\" and @value=\"OnlyIsNotVerified\"]";

        private const string ORDERS_TYPES_FILTER = "SelectedOrderTypes_ms";
        private const string UNSELECT_ORDERS_TYPES = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SEARCH_ORDERS_TYPES = "/html/body/div[12]/div/div/label/input";
        //filter date flights 
        private const string DATE_FROM = "//*[@id=\"start-date-picker\"]";
        private const string DATE_TO = "//*[@id=\"end-date-picker\"]";

        private const string FLIGHTDATE_FROM = "flight-start-date-picker";
        private const string FLIGHTDATE_TO = "flight-end-date-picker";
        private const string TODAY = "filter-show-date-day-checkbox";
        private const string TOMORROW = "filter-show-date-tomorrow-checkbox";



        private const string VERIFIED_ICON = "/html/body/div[2]/div/div[2]/div/table/tbody/tr[*]/td[3]/div/span";
        private const string CUSTOMER_ORDER_TYPE = "dropdown-order-type";
        private const string CUSTOMER_ORDER_TYPE_BY_NUMBER = "//tr[contains(td[5], '{0}')]//td[6]";

        private const string MESSAGE_INDEX_CUSTOMER_ORDER_NOT_VIEW = "//*/span[contains(@class,'fa-solid fa-comments') and contains(@style,'color:green')]";
        private const string MESSAGE_INDEX_CUSTOMER_ORDER_IS_VIEW = "//*/span[contains(@class,'fa-solid fa-comments') and contains(@style,'color:dimgray')]";
        private const string ITEM_NAME_DETAIL = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[2]/span/span/input[2]";

        private const string FIRST_PRICE_CURRENCY = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[13]";
        private const string CREATE_FROM = "production-order-createfrombtn";
        private const string CREATE_FROM_ISPATCH = "production-order-createfrombtn";
        private const string FIRSTLINE = "//*[@id=\"tableListMenu\"]/tbody/tr[1]";
        private const string FIRSTLINENAME = "//*[@id=\"tableListMenu\"]/tbody/tr/td[5]";
        private const string GENERATLINFORMATION = "hrefTabContentInformations";
        private const string CLICK_STATUS = "/html/body/div[2]/div/div[2]/div/table/tbody/tr/td[16]/select";
        private const string PAGE_2 = "//*[@id=\"tabContentItemContainer\"]/nav/ul/li[4]/a";
        private const string LINE_CO = "/html/body/div[2]/div/div[2]/div/table/tbody/tr[*]/td[3]";
        private const string TOTAL_NUMBER = "//*/span[@class='counter']";
        private const string FILTRED_ORDER = "//*[@id='collapseIsPreorderFilter']/div/div/div[2]";

        //__________________________________Variables______________________________________

        // General

        private IWebElement _label5;
        [FindsBy(How = How.XPath, Using = COUNTER)]
        private IWebElement _counter;
        [FindsBy(How = How.Id, Using = GENERATLINFORMATION)]
        private IWebElement _generalinformation;
        [FindsBy(How = How.XPath, Using = FIRSTLINENAME)]
        private IWebElement _getfirstlinename;
        [FindsBy(How = How.XPath, Using = FIRSTLINE)]
        private IWebElement _firstline;
        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusButton;

        [FindsBy(How = How.XPath, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.Id, Using = PRINT_CUSTOMER_ORDER)]
        private IWebElement _printCustomerOrder;

        [FindsBy(How = How.XPath, Using = PAGE_2)]
        private IWebElement _secondPage;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = SEND_MAIL_BTN)]
        private IWebElement _sendByMailBtn;

        [FindsBy(How = How.Id, Using = SEND)]
        private IWebElement _sendBtn;


        // Tableau customer order
        [FindsBy(How = How.Id, Using = NEW_CUSTOMER_ORDER)]
        private IWebElement _createNewCustomerOrder;

        [FindsBy(How = How.XPath, Using = FIRST_CUSTOMER_ORDER_NUMBER)]
        private IWebElement _firstCustomerOrderNumber;

        [FindsBy(How = How.XPath, Using = DELETE_BTN)]
        private IWebElement _deletebtn;

        [FindsBy(How = How.XPath, Using = DELETE_CONFIRM)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = FIRST_ORDER_NUMBER)]
        private IWebElement _firstOrderNumber;

        [FindsBy(How = How.XPath, Using = FIRST_CUSTOMER)]
        private IWebElement _firstCustomer;

        [FindsBy(How = How.XPath, Using = FIRST_DELIVERY)]
        private IWebElement _firstDelivery;

        [FindsBy(How = How.XPath, Using = VERIFIED_ALL)]
        private IWebElement _isVerifiedAll;

        [FindsBy(How = How.XPath, Using = VERIFIED_ONLY)]
        private IWebElement _isOnlyVerified;

        [FindsBy(How = How.XPath, Using = VERIFIED_NOTONLY)]
        private IWebElement _isOnlyNotVerified;

        [FindsBy(How = How.XPath, Using = FIRST_PRICE_CURRENCY)]
        private IWebElement _FirstPriceCurrency;

        [FindsBy(How = How.ClassName, Using = "counter")]
        private IWebElement _totalNumber;

        // ____________________________________________ Filtres ________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortByFilter;

        [FindsBy(How = How.Id, Using = FROM_FILTER)]
        private IWebElement _fromFilter;

        [FindsBy(How = How.Id, Using = TO_FILTER)]
        private IWebElement _toFilter;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_SITE)]
        private IWebElement _unselectSite;

        [FindsBy(How = How.Id, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.XPath, Using = SHOW_ALL_ORDER)]
        private IWebElement _showAllOrders;

        [FindsBy(How = How.XPath, Using = SHOW_INVOICED_ONLY)]
        private IWebElement _showInvoicedOnly;

        [FindsBy(How = How.XPath, Using = SHOW_NOT_INVOICED_ONLY)]
        private IWebElement _showNotInvoicedOnly;

        [FindsBy(How = How.XPath, Using = SHOW_STATUS)]
        private IWebElement _showStatus;

        [FindsBy(How = How.XPath, Using = STATUS_ALL)]
        private IWebElement _statusAll;

        [FindsBy(How = How.XPath, Using = STATUS_OPENED)]
        private IWebElement _statusOpened;

        [FindsBy(How = How.XPath, Using = STATUS_CLOSED)]
        private IWebElement _statusClosed;

        [FindsBy(How = How.XPath, Using = SHOW_INVOICE_STEP)]
        private IWebElement _showInvoiceStep;

        [FindsBy(How = How.Id, Using = INVOICE_ALL)]
        private IWebElement _invoiceAll;

        [FindsBy(How = How.Id, Using = INVOICE_IN_PROGRESS)]
        private IWebElement _invoiceInProgress;

        [FindsBy(How = How.Id, Using = INVOICE_VALIDATED)]
        private IWebElement _invoiceValidate;

        [FindsBy(How = How.Id, Using = INVOICE_INVOICED)]
        private IWebElement _invoiceInvoiced;

        [FindsBy(How = How.XPath, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_CUSTOMERS)]
        private IWebElement _unselectAllCustomers;

        [FindsBy(How = How.XPath, Using = CUSTOMER_SEARCH)]
        private IWebElement _customerSearch;

        [FindsBy(How = How.Id, Using = FLIGHT_NUMBER_FILTER)]
        private IWebElement _flightNumberFilter;

        [FindsBy(How = How.Id, Using = ORDERS_TYPES_FILTER)]
        private IWebElement _ordersTypeFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ORDERS_TYPES)]
        private IWebElement _unselectordersType;

        [FindsBy(How = How.Id, Using = SEARCH_ORDERS_TYPES)]
        private IWebElement _searchordersType;

        [FindsBy(How = How.XPath, Using = DATE_FROM)]
        private IWebElement _dateFrom;

        [FindsBy(How = How.XPath, Using = DATE_TO)]
        private IWebElement _dateTo;

        [FindsBy(How = How.XPath, Using = ITEM_NAME_DETAIL)]
        private IWebElement _itemName;

        [FindsBy(How = How.Id, Using = TOMORROW)]
        private IWebElement _tomorrow;

        [FindsBy(How = How.Id, Using = TODAY)]
        private IWebElement _today;

        [FindsBy(How = How.Id, Using = FLIGHTDATE_TO)]
        private IWebElement _flightDateTo;

        [FindsBy(How = How.Id, Using = FLIGHTDATE_FROM)]
        private IWebElement _flightDateFrom;

        public int GetCustomerOrdersCount()
        {
            WaitPageLoading();
            _counter = WaitForElementExists(By.XPath(COUNTER));
            return int.Parse(_counter.Text);
        }

        public enum FilterType
        {
            Search,
            SortBy,
            From,
            To,
            Site,
            ShowAllOrders,
            ShowInvoicedOnly,
            ShowNotInvoicedOnly,
            StatusAll,
            StatusOpened,
            StatusClosed,
            InvoiceAll,
            InvoiceInProgress,
            InvoiceValidate,
            InvoiceInvoiced,
            Customers,
            PreOrderYes,
            PreOrderNo,
            PreOrderYesNo,
            PreOrderNoNo,
            IsAllDelivered,
            IsDelivered,
            IsNotDelivered,
            IsVerified,
            IsAllVerified,
            IsNotVerified,
            FlightNumber,
            CrewPreOrder,
            OrdersTypes,
            DateFrom,
            DateTo,
            FlightDateTo,
            FlightDateFrom,
            Today,
            Tomorrow
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);
            IWebElement preOrderCheckBoxYES;
            IWebElement preOrderCheckBoxNO;

            switch (filterType)
            {
                case FilterType.Search:
                    action = new Actions(_webDriver);
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    action.MoveToElement(_searchFilter).Perform();
                    _searchFilter.SetValue(ControlType.TextBox, value);

                    WaitForLoad();
                    break;
                case FilterType.SortBy:
                    _sortByFilter = WaitForElementIsVisible(By.Id(SORTBY_FILTER));
                    _sortByFilter.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortByFilter.SetValue(ControlType.DropDownList, element.Text);
                    _sortByFilter.Click();
                    break;
                case FilterType.From:
                    _fromFilter = WaitForElementIsVisible(By.Id(FROM_FILTER));
                    _fromFilter.SetValue(ControlType.DateTime, value);
                    _fromFilter.SendKeys(Keys.Tab);
                    break;
                case FilterType.To:
                    _toFilter = WaitForElementIsVisible(By.Id(TO_FILTER));
                    _toFilter.SetValue(ControlType.DateTime, value);
                    _toFilter.SendKeys(Keys.Tab);
                    break;
                case FilterType.Site:
                    _siteFilter = WaitForElementIsVisible(By.Id(SITE_FILTER));
                    _siteFilter.Click();

                    // On décoche toutes les options
                    _unselectSite = _webDriver.FindElement(By.XPath(UNSELECT_SITE));
                    _unselectSite.Click();

                    var siteSelected = value + " - " + value;
                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                    _searchSite.SetValue(ControlType.TextBox, siteSelected);
                    WaitForLoad();

                    var valueToSite = WaitForElementIsVisible(By.XPath("//span[text()='" + siteSelected + "']"));
                    valueToSite.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();

                    _siteFilter.Click();
                    break;
                case FilterType.ShowAllOrders:
                    WaitForElementExists(By.XPath(SHOW_ALL_ORDER));
                    _showAllOrders.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowInvoicedOnly:
                    WaitForElementExists(By.XPath(SHOW_INVOICED_ONLY));
                    _showInvoicedOnly.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowNotInvoicedOnly:
                    WaitForElementExists(By.XPath(SHOW_NOT_INVOICED_ONLY));
                    _showNotInvoicedOnly.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.StatusAll:
                    WaitForElementExists(By.XPath(STATUS_ALL));
                    _statusAll.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.StatusOpened:
                    WaitForElementExists(By.XPath(STATUS_OPENED));
                    _statusOpened.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.StatusClosed:
                    WaitForElementExists(By.XPath(STATUS_CLOSED));
                    action.MoveToElement(_statusClosed).Perform();
                    _statusClosed.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();
                    break;
                case FilterType.InvoiceAll:
                    WaitForElementExists(By.Id(INVOICE_ALL));
                    action.MoveToElement(_invoiceAll).Perform();
                    _invoiceAll.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.InvoiceInProgress:
                    WaitForElementExists(By.Id(INVOICE_IN_PROGRESS));
                    action.MoveToElement(_invoiceInProgress).Perform();
                    _invoiceInProgress.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.InvoiceValidate:
                    var collapsed = WaitForElementIsVisible(By.XPath("//*/a[@href='#collapseInvoiceStepFilter']"));
                    if (collapsed.GetAttribute("class").Contains("collapse"))
                    {
                        collapsed.Click();
                    }
                    WaitForElementExists(By.Id(INVOICE_VALIDATED));
                    action.MoveToElement(_invoiceValidate).Perform();
                    _invoiceValidate.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.InvoiceInvoiced:
                    WaitForElementExists(By.Id(INVOICE_INVOICED));
                    action.MoveToElement(_invoiceInvoiced).Perform();
                    _invoiceInvoiced.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Customers:
                    var _invoiceCustomer = WaitForElementExists(By.Id(CUSTOMER_FILTER));
                    action.MoveToElement(_invoiceCustomer).Perform();
                    ComboBoxSelectById(new ComboBoxOptions(CUSTOMER_FILTER, (string)value));
                    break;
                case FilterType.PreOrderYes:
                    OpenCollapsePreOrder();

                    preOrderCheckBoxYES = WaitForElementExists(By.Id("filter-show-is-preorders-checkbox"));
                    new Actions(_webDriver).MoveToElement(preOrderCheckBoxYES).Perform();
                    preOrderCheckBoxYES.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    preOrderCheckBoxNO = WaitForElementExists(By.Id("filter-show-is-not-preorders-checkbox"));
                    new Actions(_webDriver).MoveToElement(preOrderCheckBoxNO).Perform();
                    preOrderCheckBoxNO.SetValue(ControlType.CheckBox, false);
                    WaitForLoad();
                    break;

                case FilterType.PreOrderNo:
                    OpenCollapsePreOrder();

                    preOrderCheckBoxYES = WaitForElementExists(By.Id("filter-show-is-preorders-checkbox"));
                    new Actions(_webDriver).MoveToElement(preOrderCheckBoxYES).Perform();
                    preOrderCheckBoxYES.SetValue(ControlType.CheckBox, false);
                    WaitForLoad();
                    preOrderCheckBoxNO = WaitForElementExists(By.Id("filter-show-is-not-preorders-checkbox"));
                    new Actions(_webDriver).MoveToElement(preOrderCheckBoxNO).Perform();
                    preOrderCheckBoxNO.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;

                case FilterType.PreOrderYesNo:
                    OpenCollapsePreOrder();

                    preOrderCheckBoxYES = WaitForElementExists(By.Id("filter-show-is-preorders-checkbox"));
                    new Actions(_webDriver).MoveToElement(preOrderCheckBoxYES).Perform();
                    WaitForLoad();
                    preOrderCheckBoxYES.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    preOrderCheckBoxNO = WaitForElementExists(By.Id("filter-show-is-not-preorders-checkbox"));
                    new Actions(_webDriver).MoveToElement(preOrderCheckBoxNO).Perform();
                    WaitForLoad();
                    preOrderCheckBoxNO.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;

                case FilterType.PreOrderNoNo:
                    OpenCollapsePreOrder();

                    preOrderCheckBoxYES = WaitForElementExists(By.Id("filter-show-is-preorders-checkbox"));
                    new Actions(_webDriver).MoveToElement(preOrderCheckBoxYES).Perform();
                    preOrderCheckBoxYES.SetValue(ControlType.CheckBox, false);
                    WaitForLoad();
                    preOrderCheckBoxNO = WaitForElementExists(By.Id("filter-show-is-not-preorders-checkbox"));
                    new Actions(_webDriver).MoveToElement(preOrderCheckBoxNO).Perform();
                    preOrderCheckBoxNO.SetValue(ControlType.CheckBox, false);
                    WaitForLoad();
                    break;

                case FilterType.IsAllDelivered:
                    OpenCollapseIsDelivery();

                    action = new Actions(_webDriver);
                    var isDeliveredAll = WaitForElementExists(By.XPath("//*[@id=\"DeliveredStatus\" and @value=\"All\"]"));
                    action.MoveToElement(isDeliveredAll).Perform();
                    isDeliveredAll.Click();
                    
                    break;

                case FilterType.IsDelivered:
                    OpenCollapseIsDelivery();
                   
                    action = new Actions(_webDriver);
                    var isOnlyDelivered = WaitForElementExists(By.XPath("//*[@id=\"DeliveredStatus\" and @value=\"Yes\"]"));
                    action.MoveToElement(isOnlyDelivered).Perform();
                    isOnlyDelivered.Click();
                    
                    break;

                case FilterType.IsNotDelivered:
                    OpenCollapseIsDelivery();
                    
                    action = new Actions(_webDriver);
                    var isOnlyNotDelivered = WaitForElementExists(By.XPath("//*[@id=\"DeliveredStatus\" and @value=\"No\"]"));
                    action.MoveToElement(isOnlyNotDelivered).Perform();
                    isOnlyNotDelivered.Click();

                    break;

                case FilterType.IsVerified:
                    OpenCollapseIsVerified();

                    ScrollUntilElementIsVisible(By.XPath(VERIFIED_ONLY));
                    action = new Actions(_webDriver);
                    _isOnlyVerified = WaitForElementExists(By.XPath(VERIFIED_ONLY));
                    action.MoveToElement(_isOnlyVerified).Perform();
                    _isOnlyVerified.Click();
                    
                    break;

                case FilterType.IsNotVerified:
                    OpenCollapseIsVerified();

                    ScrollUntilElementIsVisible(By.XPath(VERIFIED_NOTONLY));
                    action = new Actions(_webDriver);
                    _isOnlyNotVerified = WaitForElementExists(By.XPath(VERIFIED_NOTONLY));
                    action.MoveToElement(_isOnlyNotVerified).Perform();
                    _isOnlyNotVerified.Click();
                   
                    break;

                case FilterType.IsAllVerified:
                    OpenCollapseIsVerified();

                    action = new Actions(_webDriver);
                    _isVerifiedAll = WaitForElementExists(By.XPath(VERIFIED_ALL));
                    action.MoveToElement(_isVerifiedAll).Perform();
                    _isVerifiedAll.Click();
                    
                    break;

                case FilterType.FlightNumber:
                    if (isElementExists(By.XPath("//*[@id=\"item-filter-form\"]/div[14]/a[@class=\"filterCollapseButton collapsed\"]")))
                    {
                        ScrollUntilElementIsVisible(By.XPath("//*[@id=\"item-filter-form\"]/div[14]/a[@class=\"filterCollapseButton collapsed\"]"));
                        WaitForElementExists(By.XPath("//*[@id=\"item-filter-form\"]/div[14]/a[@class=\"filterCollapseButton collapsed\"]")).Click();
                    }
                    ScrollUntilElementIsVisible(By.Id(FLIGHT_NUMBER_FILTER));
                    _flightNumberFilter = WaitForElementExists(By.Id(FLIGHT_NUMBER_FILTER));
                    _flightNumberFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;

                case FilterType.CrewPreOrder:

                    ScrollUntilElementIsVisible(By.XPath("//*[@id=\"item-filter-form\"]/div[13]/a[@class=\"filterCollapseButton collapsed\"]"));
                    if (isElementExists(By.XPath("//*[@id=\"item-filter-form\"]/div[13]/a[@class=\"filterCollapseButton collapsed\"]")))
                    {
                        WaitForElementExists(By.XPath("//*[@id=\"item-filter-form\"]/div[13]/a[@class=\"filterCollapseButton collapsed\"]")).Click();
                    }
                    if ((bool)value)
                    {
                        ScrollUntilElementIsVisible(By.XPath("//*[@id=\"collapseIsCrewPreorderFilter\"]/div/div/div[2]/div/div"));

                        var CrewPreOrderCheckBoxNo = WaitForElementExists(By.XPath("//*[@id=\"collapseIsCrewPreorderFilter\"]/div/div/div[2]/div/div"));
                        CrewPreOrderCheckBoxNo.Click();
                    }
                    else
                    {
                        ScrollUntilElementIsVisible(By.XPath("//*[@id=\"collapseIsCrewPreorderFilter\"]/div/div/div[1]/div/div"));
                        var CrewPreOrderCheckBoxYES = WaitForElementExists(By.XPath("//*[@id=\"collapseIsCrewPreorderFilter\"]/div/div/div[1]/div/div"));
                        CrewPreOrderCheckBoxYES.Click();
                    }
                    break;

                case FilterType.OrdersTypes:
                    _ordersTypeFilter = WaitForElementIsVisible(By.Id(ORDERS_TYPES_FILTER));
                    _ordersTypeFilter.Click();

                    // On décoche toutes les options
                    _unselectordersType = _webDriver.FindElement(By.XPath(UNSELECT_ORDERS_TYPES));
                    _unselectordersType.Click();

                    var OrdersTypesSelected = value;
                    _searchordersType = WaitForElementIsVisible(By.XPath(SEARCH_ORDERS_TYPES));
                    _searchordersType.SetValue(ControlType.TextBox, OrdersTypesSelected);
                    WaitForLoad();

                    var valueToOrdersTypes = WaitForElementIsVisible(By.XPath("//span[text()='" + OrdersTypesSelected + "']"));
                    valueToOrdersTypes.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();

                    _ordersTypeFilter.Click();
                    break;

                case FilterType.DateFrom:
                    _dateFrom = WaitForElementIsVisible(By.XPath(DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);

                    _dateFrom.SendKeys(Keys.Tab);
                    WaitForLoad();
                    break;

                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisible(By.XPath(DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);

                    _dateTo.SendKeys(Keys.Tab);
                    WaitForLoad();
                    break;
                
                case FilterType.Today:
                    _today = WaitForElementExists(By.Id(TODAY));
                    action.MoveToElement(_today).Perform();
                    _today.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;

                case FilterType.Tomorrow:
                    _tomorrow = WaitForElementExists(By.Id(TOMORROW));
                    action.MoveToElement(_tomorrow).Perform();
                    _tomorrow.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;

                case FilterType.FlightDateFrom:
                    _flightDateFrom = WaitForElementExists(By.Id(FLIGHTDATE_FROM));
                    action.MoveToElement(_flightDateFrom).Perform();
                    _flightDateFrom.SetValue(ControlType.DateTime, value);
                    _flightDateFrom.SendKeys(Keys.Tab);
                    WaitForLoad();
                    break;

                case FilterType.FlightDateTo:
                    _flightDateTo = WaitForElementExists(By.Id(FLIGHTDATE_TO));
                    action.MoveToElement(_flightDateTo).Perform();
                    _flightDateTo.SetValue(ControlType.DateTime, value);
                    _flightDateTo.SendKeys(Keys.Tab);
                    WaitForLoad();
                    break;

                default:
                    break;
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void OpenCollapsePreOrder()
        {
            if (isElementExists(By.XPath("//*/label[text()='Preorders']/../../a[contains(@class,'collapsed')]")))
            {
                var collapse = WaitForElementExists(By.XPath("//*/label[text()='Preorders']/../../a[contains(@class,'collapsed')]"));
                new Actions(_webDriver).MoveToElement(collapse).Click().Perform();
                WaitPageLoading();
                WaitForLoad();
            }
        }

        public void OpenCollapseIsDelivery()
        {
            if (isElementExists(By.XPath("//*/label[text()='Is Delivered ']/../../a[contains(@class,'collapsed')]")))
            {
                var collapse = WaitForElementExists(By.XPath("//*/label[text()='Is Delivered ']/../../a[contains(@class,'collapsed')]"));
                new Actions(_webDriver).MoveToElement(collapse).Click().Perform();
                WaitPageLoading();
                WaitForLoad();
            }
        }

        public void OpenCollapseIsVerified()
        {
            if (isElementExists(By.XPath("//*/label[text()='IsVerified ']/../../a[contains(@class,'collapsed')]")))
            {
                var collapse = WaitForElementExists(By.XPath("//*/label[text()='IsVerified ']/../../a[contains(@class,'collapsed')]"));
                new Actions(_webDriver).MoveToElement(collapse).Click().Perform();
                WaitPageLoading();
                WaitForLoad();
            }
        }

        public void ResetFilter()
        {
            var html = _webDriver.FindElement(By.TagName("html"));
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);

            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);

            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(FilterType.To, DateUtils.Now);
            }

        }

        public void ClickInvoiceSteps()
        {
            _showInvoiceStep = WaitForElementIsVisible(By.XPath(SHOW_INVOICE_STEP));
            _showInvoiceStep.Click();

            // Temps que les boutons soient accessibles
            WaitForLoad();
        }
        public CustomerOrderItem ClickFirstLine()
        {
            WaitForLoad();
            var line = WaitForElementIsVisible(By.XPath(FIRST_ORDER_NUMBER));
            line.Click();
            WaitForLoad();
            return new CustomerOrderItem(_webDriver, _testContext);
        }

        public void ClickGeneralInformation()
        {
            _generalinformation = WaitForElementIsVisible(By.Id(GENERATLINFORMATION));
            _generalinformation.Click();
            WaitForLoad();
        }

        public void ClickStatus()
        {
            _showStatus = WaitForElementIsVisible(By.XPath(SHOW_STATUS));
            _showStatus.Click();

            // Temps que les boutons soient accessibles
            WaitForLoad();
            Thread.Sleep(2000);
        }

        public bool CheckValidation(bool validated)
        {
            bool isValidated = true;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[1]/div/i[contains(@class,'fa-regular fa-circle-check')]", i + 1)));

                    if (!validated)
                        return true;
                }
                catch
                {
                    isValidated = false;

                    if (validated)
                        return false;
                }
            }

            return isValidated;
        }

        public bool IsInvoiced()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    var element = _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[2]/div/i", i + 1)));

                    if (!element.GetAttribute("title").Equals("Fully or partially invoiced"))
                        return false;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsNotInvoiced()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    var element = _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[2]/div/i", i + 1)));

                    if (!element.GetAttribute("title").Equals("Not invoiced yet"))
                        return false;
                }
                catch
                {
                    return false;
                }
            }
            return true;

        }

        public bool IsSortedById()
        {
            var listeId = _webDriver.FindElements(By.XPath(CUSTOMER_ORDER_NUMBER));

            if (listeId.Count == 0)
                return false;

            var ancienId = listeId[0].Text;

            foreach (var elm in listeId)
            {
                try
                {
                    if (elm.Text.CompareTo(ancienId) > 0)
                        return false;

                    ancienId = elm.Text;
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
            var customers = _webDriver.FindElements(By.XPath(CUSTOMER_NAME));

            if (customers.Count == 0)
                return false;

            foreach (var elm in customers)
            {
                if (elm.Text != customer)
                {
                    return false;
                }
            }
            return true;
        }

        public bool IsSortedByCustomer()
        {
            var listeCustomers = _webDriver.FindElements(By.XPath(CUSTOMER_NAME));

            if (listeCustomers.Count == 0)
                return false;

            var ancienCustomer = listeCustomers[0].Text;

            foreach (var elm in listeCustomers)
            {
                try
                {
                    if (elm.Text.CompareTo(ancienCustomer) < 0)
                        return false;

                    ancienCustomer = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsSortedByOrderDate(string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            ReadOnlyCollection<IWebElement> dates;
            dates = _webDriver.FindElements(By.XPath("//*[@id='tableListMenu']/tbody/tr[*]/td[11]"));

            if (dates.Count == 0)
                return false;

            var ancientDate = DateTime.Parse(dates[0].Text, ci);

            foreach (var elm in dates)
            {
                try
                {
                    if (DateTime.Compare(ancientDate.Date, DateTime.Parse(elm.Text, ci)) < 0)
                        return false;

                    ancientDate = DateTime.Parse(elm.Text, ci);
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsDateRespected(DateTime fromDate, DateTime toDate, string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            ReadOnlyCollection<IWebElement> dates;
            dates = _webDriver.FindElements(By.XPath("//*[@id='tableListMenu']/tbody/tr[*]/td[11]"));

            if (dates.Count == 0)
                return false;

            foreach (var elm in dates)
            {
                try
                {
                    DateTime date = DateTime.Parse(elm.Text, ci);

                    if (DateTime.Compare(date, fromDate.Date) < 0 || DateTime.Compare(date, toDate.Date) > 0)
                        return false;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsSentByMail()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format(SENT_BY_MAIL, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format(SENT_BY_MAIL, i + 1)));
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

        public List<string> GetCustomerOrdersIdList()
        {
            var customerOrdersIdList = new List<string>();

            var customerOrdersList = _webDriver.FindElements(By.XPath(CUSTOMER_ORDER_NUMBER));

            foreach (var customerOrder in customerOrdersList)
            {
                customerOrdersIdList.Add(customerOrder.Text.Remove(0, 3));
            }

            return customerOrdersIdList;
        }

        // ___________________________________________________ Méthodes ________________________________________________________

        // General
        public override void ShowPlusMenu()
        {
            var action = new Actions(_webDriver);

            _plusButton = WaitForElementExists(By.XPath(PLUS_BTN));
            action.MoveToElement(_plusButton).Perform();
        }

        /*public override void ShowExtendedMenu()
        {
            var action = new Actions(_webDriver);
            WaitForElementIsVisible(By.XPath(EXTENDED_BUTTON));
            action.MoveToElement(_extendedButton).Perform();
            WaitForLoad();
        }*/

        public void ExportExcel(bool versionPrint)
        {

            ShowExtendedMenu();
            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            }

            WaitForDownload();
        }

        public FileInfo GetExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
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

        public bool IsExcelFileCorrect(string filePath)
        {
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // minutes

            Regex reg = new Regex("^export-orders" + " " + annee + "-" + mois + "-" + jour + " " + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = reg.Match(filePath);

            return m.Success;
        }

        public void SendByMail()
        {
            ShowExtendedMenu();

            _sendByMailBtn = WaitForElementIsVisible(By.Id(SEND_MAIL_BTN));
            _sendByMailBtn.Click();
            WaitForLoad();

            _sendBtn = WaitForElementIsVisible(By.Id(SEND));
            _sendBtn.Click();

            // Temps de fermeture de la fenêtre

            WaitForLoad();
            WaitPageLoading();
        }


        // Tableau
        public CustomerOrderCreateModalPage CustomerOrderCreatePage()
        {
            ShowPlusMenu();

            var elementId = IsDev() ? NEW_CUSTOMER_ORDER : NEW_CUSTOMER_ORDER_ISPATCH;
            _createNewCustomerOrder = WaitForElementIsVisible(By.Id(elementId));
            _createNewCustomerOrder.Click();
            WaitForLoad();

            return new CustomerOrderCreateModalPage(_webDriver, _testContext);
        }

        public string GetCustomerOrderNumber()
        {
            _firstCustomerOrderNumber = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER_ORDER_NUMBER));
            return _firstCustomerOrderNumber.Text;
        }

        public string GetFirstLineName()
        {
            var element = WaitForElementIsVisible(By.XPath(FIRSTLINENAME));
            var text = element.Text;
            var match = Regex.Match(text, @"(?i)Id\s(\d+)");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            else
            {
                return string.Empty;
            }
        }


        public void DeleteCustomerOrder()
        {
            WaitPageLoading();
            WaitForLoad();
            _deletebtn = WaitForElementIsVisible(By.XPath(DELETE_BTN));
            _deletebtn.Click();
            WaitForLoad();

            if (isElementVisible(By.XPath(DELETE_CONFIRM)))
            {
                _confirmDelete = WaitForElementIsVisible(By.XPath(DELETE_CONFIRM));
                _confirmDelete.Click();
                WaitForLoad();
            }
            WaitPageLoading();
        }

        public CustomerOrderItem SelectFirstCustomerOrder()
        {
            _firstCustomerOrderNumber = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER_ORDER_NUMBER));
            _firstCustomerOrderNumber.Click();
            WaitPageLoading();
            WaitForLoad();

            return new CustomerOrderItem(_webDriver, _testContext);
        }
        public string GetFirstOrderNumber()
        {
            _firstOrderNumber = WaitForElementIsVisible(By.XPath(FIRST_ORDER_NUMBER));
            return _firstOrderNumber.Text;
        }

        public string GetFirstCustomer()
        {
            _firstCustomer = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER));
            return _firstCustomer.Text;
        }

        public string GetFirstDelivery()
        {
            _firstDelivery = WaitForElementIsVisible(By.XPath(FIRST_DELIVERY));
            return _firstDelivery.Text;
        }

        public PrintReportPage Print()
        {

            ShowExtendedMenu();

            _printCustomerOrder = WaitForElementIsVisible(By.Id(PRINT_CUSTOMER_ORDER));
            _printCustomerOrder.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public bool IsVerified()
        {
            ShowExtendedMenu();

            if (isElementVisible(By.XPath("//*/a[contains(@class,'btn-verify')]")))
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        public bool isOrderTypeExist(string name)
        {
            _ordersTypeFilter = WaitForElementIsVisible(By.Id(ORDERS_TYPES_FILTER));
            _ordersTypeFilter.Click();

            var listOrdersTypes = _webDriver.FindElements(By.XPath("/html/body/div[12]/ul/li[*]"));
            List<string> ordersTypes = new List<string>();
            foreach (var orderType in listOrdersTypes)
            {
                ordersTypes.Add(orderType.Text);
            }

            if (!ordersTypes.Contains(name))
            {
                return false;
            }
            return true;
        }
        public bool VerifyDate(DateTime dateFrom, DateTime dateTo)
        {
            var listCODate = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[11]"));
            List<DateTime> dates = new List<DateTime>();
            foreach (var dateElement in listCODate)
            {
                DateTime parsedDate;
                if (DateTime.TryParseExact(dateElement.Text.Substring(0, 10), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                {
                    dates.Add(parsedDate.Date);
                }
            }
            foreach (var date in dates)
            {
                if (date <= dateFrom.Date && date >= dateTo.Date)
                {
                    return false;
                }
            }
            return true;
        }
        public void ScrollUntilElementIsVisible(By by)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            var element = WaitForElementExists(by);
            if (element != null)
            {
                js.ExecuteScript("arguments[0].scrollIntoView(true);", element);
            }
        }
        public void ScrollUntilElementIsInView(By by)
        {
            bool elementFound = false;
            int retryCount = 3; // Number of retries

            while (!elementFound && retryCount > 0)
            {
                try
                {
                    var element = WaitForElementIsVisible(by);
                    var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                    javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", element);
                    elementFound = true; // Element found and scrolled successfully
                }
                catch (StaleElementReferenceException)
                {
                    retryCount--; // Retry if a stale element exception is thrown
                }
            }

            if (!elementFound)
            {
                throw new NoSuchElementException("Unable to find the element after multiple retries.");
            }
        }
        public bool VerifyAllVerified()
        {
            var iconTitles = _webDriver.FindElements(By.XPath(VERIFIED_ICON));
            if (iconTitles.Count > 0)
            {
                foreach (var icon in iconTitles)
                {
                    if (icon.GetAttribute("title") != "Verified")
                        return false;
                }
            }
            else
            {
                return false;
            }
            return true;
        }

        public bool VerifyAllNotVerified()
        {
            var iconTitles = _webDriver.FindElements(By.XPath(VERIFIED_ICON));
            if (iconTitles.Count > 0)
            {
                foreach (var icon in iconTitles)
                {
                    if (icon.GetAttribute("title") != "Not Verified")
                        return false;
                }
            }
            else
            {
                return false;
            }
            return true;
        }
        public enum CustomerOrderType
        {
            OrderTest,
            TestOrderType,
            TestAuto
        }

        public void SetOrderType(CustomerOrderType type)
        {
            var typeElement = WaitForElementIsVisible(By.Id(CUSTOMER_ORDER_TYPE));
            switch (type)
            {
                case CustomerOrderType.OrderTest:
                    typeElement.SetValue(ControlType.DropDownList, "orderTest");
                    break;
                case CustomerOrderType.TestOrderType:
                    typeElement.SetValue(ControlType.DropDownList, "Test Order Type");
                    break;
                case CustomerOrderType.TestAuto:
                    typeElement.SetValue(ControlType.DropDownList, "testAuto");
                    break;
            }
            WaitPageLoading();
            WaitForLoad();
        }



        public string GetCustomerOrderTypeByNumber(string customerOrderNumber)
        {
            var result = WaitForElementIsVisible(By.XPath(string.Format(CUSTOMER_ORDER_TYPE_BY_NUMBER, customerOrderNumber))).Text;
            return result;
        }

        public string GetFirstExpeditionDate()
        {
            var expeditionDate = WaitForElementIsVisible(By.XPath(FIRST_EXPEDITION_DATE));
            return expeditionDate.Text;
        }

        public bool VerifyGreenPictoMessageInCutomerOrderWhenNotView()
        => isElementVisible(By.XPath(MESSAGE_INDEX_CUSTOMER_ORDER_NOT_VIEW));
        public bool VerifyGrayPictoMessageInCutomerOrderWhenIsView()
        => isElementVisible(By.XPath(MESSAGE_INDEX_CUSTOMER_ORDER_IS_VIEW));

        public void ShowBtnPlus()
        {
            var _createNewRowItem = WaitForElementIsVisible(By.Id("addOrderDetailBtn"));
            _createNewRowItem.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void FillField_CreatNewItemDetailOrder(string item)
        {
            Random rnd = new Random();

            _itemName = WaitForElementExists(By.XPath(ITEM_NAME_DETAIL));
            _itemName.Click();
            _itemName.SetValue(ControlType.TextBox, item);
            _itemName.SendKeys(Keys.ArrowDown);
            _itemName.SendKeys(Keys.Enter);
            WaitForLoad();


        }
        public string GetFirstPriceCurrency()
        {

            _FirstPriceCurrency = WaitForElementExists(By.XPath(FIRST_PRICE_CURRENCY));
            var result = _FirstPriceCurrency.Text;
            return result.Substring(0, result.Length - 1).Replace(" ", "");


        }
        public CustomerOrderCreateModalPage CreateFromCreatePage()
        {
            ShowPlusMenu();

            var elementId = IsDev() ? CREATE_FROM : CREATE_FROM_ISPATCH;
            _createNewCustomerOrder = WaitForElementIsVisible(By.Id(elementId));
            _createNewCustomerOrder.Click();
            WaitForLoad();

            return new CustomerOrderCreateModalPage(_webDriver, _testContext);
        }
        public bool IsEyeIconGreenForOrder()
        {

            int expectedR = 105;
            int expectedG = 163;
            int expectedB = 5;


            var eyeIcon = WaitForElementIsVisible(By.XPath(EYE_ICON_GREEN));


            var jsExecutor = (IJavaScriptExecutor)_webDriver;
            string color = (string)jsExecutor.ExecuteScript("return window.getComputedStyle(arguments[0]).color;", eyeIcon);

            Console.WriteLine("Computed Color: " + color);

            var colorMatch = System.Text.RegularExpressions.Regex.Match(color, @"rgba?\((\d+),\s*(\d+),\s*(\d+)\)");

            if (colorMatch.Success)
            {
                int r = int.Parse(colorMatch.Groups[1].Value);
                int g = int.Parse(colorMatch.Groups[2].Value);
                int b = int.Parse(colorMatch.Groups[3].Value);

                Console.WriteLine($"Actual RGB Values (Computed): R={r}, G={g}, B={b}");

                return r == expectedR && g == expectedG && b == expectedB;
            }

            return false;
        }
        public bool IsEyeIconGrayForOrder()
        {
            int expectedR = 128;
            int expectedG = 128;
            int expectedB = 128;

            var eyeIcon = WaitForElementIsVisible(By.XPath(EYE_ICON_GRAY));

            var jsExecutor = (IJavaScriptExecutor)_webDriver;
            string color = (string)jsExecutor.ExecuteScript("return window.getComputedStyle(arguments[0]).color;", eyeIcon);

            Console.WriteLine("Computed Color: " + color);

            var colorMatch = System.Text.RegularExpressions.Regex.Match(color, @"rgba?\((\d+),\s*(\d+),\s*(\d+)\)");

            if (colorMatch.Success)
            {
                int r = int.Parse(colorMatch.Groups[1].Value);
                int g = int.Parse(colorMatch.Groups[2].Value);
                int b = int.Parse(colorMatch.Groups[3].Value);

                Console.WriteLine($"Actual RGB Values (Computed): R={r}, G={g}, B={b}");

                return r == expectedR && g == expectedG && b == expectedB;
            }

            return false;
        }

        public string GetFirstCustomerOrderNumber()
        {
            _firstOrderNumber = WaitForElementIsVisible(By.XPath(FIRST_ORDER_NUMBER));
            string fullText = _firstOrderNumber.Text;
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(fullText);
            if (match.Success)
            {
                return match.Value;
            }
            return string.Empty;
        }
        public string GetOrderNumber()
        {
            var orderElement = WaitForElementIsVisible(By.Id("production-orders-item-id-1"));
            string orderText = orderElement.Text;
            int orderNumber = int.Parse(orderText.Replace("Id ", ""));
            return orderNumber.ToString();
        }

        public CustomerOrderPage EditStatusCO(string status)
        {
            var statusSelect = WaitForElementExists(By.XPath(CLICK_STATUS));
            statusSelect.SetValue(ControlType.DropDownList, status);
            WaitForLoad();

            return new CustomerOrderPage(_webDriver, _testContext);
        }

        public bool VerifyStatus(string statusCancelled)
        {
            IWebElement dropdown = _webDriver.FindElement(By.XPath(CLICK_STATUS));

            SelectElement select = new SelectElement(dropdown);

            string selectedOptionText = select.SelectedOption.Text;

            if ((selectedOptionText != string.Empty) && (selectedOptionText.Contains(statusCancelled)))
            {
                return true;

            }
            else return false; 

        }

        public void CloseDatePicker()
        {
            _toFilter = WaitForElementIsVisible(By.Id(TO_FILTER));
            _toFilter.SendKeys(Keys.Tab);
        }
        public void SecondPage()
        {
            _secondPage = WaitForElementIsVisible(By.XPath(PAGE_2));
            _secondPage.Click();
            WaitLoading();
        }
        public int GetNumberCO()
        {
            return _webDriver.FindElements(By.XPath(LINE_CO)).Count;
        }
      
        public bool verifTomorrowFilter(string date)
        {
            _flightDateFrom = WaitForElementExists(By.Id(FLIGHTDATE_FROM));
            _flightDateTo = WaitForElementExists(By.Id(FLIGHTDATE_TO));
            string result1 = _flightDateFrom.GetAttribute("value");
            string result2 = _flightDateTo.GetAttribute("value");
            var result = result1 == date && result2 == date;
            return result;
        }
        public int CheckTotalNumber_customerOrders()
        {
            WaitForLoad();
            _totalNumber = WaitForElementExists(By.XPath(TOTAL_NUMBER));
            int nombre = Int32.Parse(_totalNumber.GetAttribute("innerText"));
            return nombre;
        }
        public List<string> GetFilteredOrders()
        {
            var orderRows = _webDriver.FindElements(By.XPath(FILTRED_ORDER));
            return orderRows.Select(row => row.Text.Trim()).ToList();
        }

        public void PageUpCO()
        {
            // Modifiez la valeur 'top' et déclenchez un événement
            var scrollbarElement = _webDriver.FindElement(By.ClassName("ps__scrollbar-y-rail"));

            // Ajoutez une interaction de glisser-déposer
            Actions actions = new Actions(_webDriver);
            actions.ClickAndHold(scrollbarElement).MoveByOffset(0, -300).Release().Perform();
        }
        public string GetQuantity()
        {
            var quantityElement = _webDriver.FindElement(By.XPath("//*[contains(@id,'DeliveryQuantity')]\r\n "));
            return quantityElement.GetAttribute("value");
        }

        public string GetUnitPrice()
        {
            var unitPriceElement = _webDriver.FindElement(By.XPath("//*[contains(@id,'UnitPrice')]\r\n "));
            return unitPriceElement.GetAttribute("value");
        }

        public string GetTotalPrice()
        {
            var totalPriceElement = _webDriver.FindElement(By.XPath("//*[contains(@id,'TotalPrice')]\r\n "));
            return totalPriceElement.GetAttribute("value");
        }

        public void SetQuantity(string newQuantity)
        {
            var quantityElement = _webDriver.FindElement(By.XPath("//*[contains(@id,'DeliveryQuantity')]\r\n "));

            quantityElement.Clear();
            // Saisir la nouvelle valeur
            quantityElement.SendKeys(newQuantity);
        }


    }
}
