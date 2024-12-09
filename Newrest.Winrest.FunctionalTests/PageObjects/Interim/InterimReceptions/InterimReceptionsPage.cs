using Limilabs.Client.IMAP;
using iText.Commons.Utils;
using Limilabs.Client.IMAP;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Security.Policy;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimReceptions
{
    public class InterimReceptionsPage : PageBase
    {
        //-----------Table Liste------------
        private const string ROWS_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr[*]";
        private const string COL_CONVERTER_TYPE = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[8]";
        private const string RESULTS_ROWS = "/html/body/div[2]/div/div[2]/div[2]/table";
        private const string COL_CUSTOMERS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[6]";
        private const string COL_IS_ACTIVE = "//*[@id=\"item_IsActive\"]";
        private const string PROVIDERS_NAMES = "//*[@id=\"tableListMenu\"]/thead/tr/th[2]";

        // Filters
        private const string RESET_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[1]/a";
        private const string SITES_FILTER = "cbSortBy";
        private const string FILTER_SEARCH_NUMBER = "SearchNumber";
        private const string FILTER_SEARC_NUMBER_SECONDARY = "SearchNumberSecondary";
        private const string FILTER_DATE_FROM = "date-picker-start";
        private const string FILTER_DATE_TO = "date-picker-end";
        private const string FILTER_SUPPLIER = "SelectedSuppliers_ms";
        private const string UNCHECK_ALL_SUPPLIER = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SUPPLIER_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string SHOW_ALL = "//*[@id=\"ShowActive\"][@value='All']";
        private const string SHOW_ONLY_ACTIVE = "//*[@id=\"ShowActive\"][@value='ActiveOnly']";
        private const string SHOW_ONLY_INACTIVE = "//*[@id=\"ShowActive\"][@value='InactiveOnly']";
        private const string SHOW_ALL_RECEPTIONS = "//*[@id=\"ShowOnlyValue\"][@value='All']";
        private const string SHOW_NOT_VALIDATED_ONLY = "//*[@id=\"ShowOnlyValue\"][@value='NotValidatedOnly']";
        private const string SHOW_VALIDATED_ONLY = "//*[@id=\"ShowOnlyValue\"][@value='ValidatedOnly']";
        private const string FILTER_INTERIM_RECEPTION_STATUS_ALL = "//*[@id=\"Status\"][@value='All']";
        private const string FILTER_INTERIM_RECEPTION_STATUS_OPENED = "//*[@id=\"Status\"][@value='Opened']";
        private const string FILTER_INTERIM_RECEPTION_STATUS_CLOSED = "//*[@id=\"Status\"][@value='Closed']";
        private const string FILTER_SORTBY = "cbSortBy";
        private const string FILTER_SITES = "SelectedSites_ms";
        private const string SEARCH_SITES = "/html/body/div[11]/div/div/label/input";
        private const string UNCHECK_ALL_SITES = "/html/body/div[11]/div/ul/li[2]/a";
        private const string LIST_SITE_FILTER = "/html/body/div[11]/ul/li[*]/label/input";
        private const string LIST_SUPPLIER_FILTER = "/html/body/div[10]/ul/li[*]/label/input";
        private const string LIST_DELIVERY_DATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[5]";
        private const string CLICKSHOWMENU = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/button";
        private const string EXPORT_BUTTON = "btn-export-excel";
        private const string EXPORT = "btn-export-excel";
        private const string FIRST_INTERIM_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[2]";
        private const string DETAIL_BUTTON = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[4]";
        private const string SHOW_SELECTED = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[7]/label[1]/input";
        private const string SUPPLIERS_SELECTED = "/html/body/div[10]/ul/li/label/input";
        private const string SITES_SELECTED = "/html/body/div[11]/ul/li/label/input";
        private const string SHOW_ACTIVE_SELECTED = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[8]/input";//
        private const string SHOW_STATUS_SELECTED = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[10]/div/input";
        private const string ERROR_MESSAGE = "//*[@id=\"dataAlertModal\"]/div/div/div[2]/p";
        private const string FIRST_INTERIM_DATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[5]";
        //Print  
        private const string PRINTINTERIMRECEPTION = "//*[@id=\"btn-print-receptions\"]";
        private const string PRINT_BTN = "//*[@id=\"printButton\"]";
        private const string ISPRICEINCLUDE = "//*[@id=\"IsPricesIncluded\"]";

        public InterimReceptionsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        //__________________________________Constantes_______________________________________

        // General
        private const string VALIDATE_RESULTS = "btn-validate-all";
        private const string CONFIRM_VALIDATE_RESULTS = "btn-popup-validate-all";
        private const string NEW_INTERIM_BUTTON = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[2]/div/a[text()='New interim reception']";
        private const string FIRST_INTERIM = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[2]";
        private const string DELETE_INTERIM_BUTTON_DEV = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[contains(text(), '{0}')]/..//span[contains(@class, 'fas fa-trash-alt')]";
        private const string DELETE_INTERIM_CONFIRM_BUTTON = "dataConfirmOK";
        private const string SUPPLIER_VALIDATOR_MESSAGE = "//*[@id=\"form-create-interim-reception\"]/div/div[3]/div[1]/div/div/span]";
        private const string DELIVERY_VALIDATOR_MESSAGE = "//*[@id=\"form-create-interim-reception\"]/div/div[4]/div[1]/div/div/span]";
        private const string LOCATION_VALIDATOR_MESSAGE = "//*[@id=\"form-create-interim-reception\"]/div/div[3]/div[2]/div/div/span]";
        private const string NEW_Interim = "New interim reception";
        private const string Interim_site = "//*[@id=\"list-item-with-action\"]/table/thead/tr/th[3]";
        private const string Byinterimordernumber_FILTER = "SearchNumberSecondary";
        private const string Suppliers_FILTER = "SelectedSuppliers_ms";
        private const string SHOW_ALL_RECEPTIONS_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[7]/label[1]/input";
        private const string SHOW_VALIDATE_ONLY_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[7]/label[2]/input";
        private const string SHOW_NOT_VALIDATE_ONLY_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[7]/label[3]/input";
        private const string SHOW_ALL_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[7]/label[1]/input";
        private const string SHOW_ACTIVE_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[7]/label[2]/input";
        private const string SHOW_INACTIVE_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[7]/label[3]/input";
        private const string BYNUMBERFilter_FILTER = "SearchNumber";
        private const string SORTBY_FILTER = "cbSortBy";
        private const string ALL_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[10]/div/input[1]";
        private const string OPENED_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[10]/div/input[2]";
        private const string CLOSED_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[10]/div/input[3]";
        private const string Sites_FILTER = "SelectedSites_ms";
        private const string SEARCH_SITE = "/html/body/div[11]/div/div/label/input";
        private const string SITES_SELECT_ALL = "/html/body/div[11]/div/ul/li[1]/a/span[2]";
        private const string SITE_UNSELECT_ALL = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string SITE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[3]";
        private const string DATES = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[5]";
        private const string INTERIM_RECEPTION_TABLE = "//*[@id=\"list-item-with-action\"]/table";
        private const string INTERIM_RECEPTION_COUNTER = "//*[@id=\"div-body\"]/div/div[2]/div[1]/h1/span";
        private const string SUPPLIERS = "SelectedSuppliers_ms";
        private const string SUPPLIER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[4]";
        private const string FIRST_LINE_NOT_VALIDATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]";
        private const string ADD_NEW_INTERIM_RECEPTION_BTN = "//*[@id=\"btn-tab-interim-reception\"]";
        private const string EXTRACTED_NUMBER_ELEMENT = "//*[@id=\"table-copy-from-ir\"]/tbody/tr[2]/td[2]";
        private const string FILTER_INPUT = "//*[@id=\"form-copy-from-tbSearchName\"]";
        private const string FILTERED_ROWS = "//*[@id=\"table-copy-from-ir\"]/tbody/tr";
        private const string FIRST_ROW_CELL = "//*[@id=\"table-copy-from-ir\"]/tbody/tr[2]/td[2]";

        private const string ALL_RN_ROWS = "//*[@id='list-item-with-action']/table/tbody/tr[*]";

        //__________________________________Variables_______________________________________

        // Général

        [FindsBy(How = How.Id, Using = ISPRICEINCLUDE)]
        private IWebElement _ispriceinclude;
        [FindsBy(How = How.Id, Using = PRINT_BTN)]
        private IWebElement _print_btn;

        [FindsBy(How = How.Id, Using = PRINTINTERIMRECEPTION)]
        private IWebElement _printinterimreception;
        [FindsBy(How = How.Id, Using = BYNUMBERFilter_FILTER)]
        private IWebElement _BynumberFilter;
        [FindsBy(How = How.Id, Using = Byinterimordernumber_FILTER)]
        private IWebElement _ByinterimordernumberFilter;
        [FindsBy(How = How.Id, Using = Suppliers_FILTER)]
        private IWebElement _SuppliersFilter;
        [FindsBy(How = How.XPath, Using = SHOW_ALL_RECEPTIONS_FILTER)]
        private IWebElement _showallreceptions;
        [FindsBy(How = How.XPath, Using = SHOW_VALIDATE_ONLY_FILTER)]
        private IWebElement _shownotvalidatedonly;
        [FindsBy(How = How.XPath, Using = SHOW_ALL_FILTER)]
        private IWebElement _showall;
        [FindsBy(How = How.XPath, Using = SHOW_ACTIVE_FILTER)]
        private IWebElement _showactive;
        [FindsBy(How = How.XPath, Using = SHOW_INACTIVE_FILTER)]
        private IWebElement _showinactive;
        [FindsBy(How = How.XPath, Using = SHOW_VALIDATE_ONLY_FILTER)]
        private IWebElement _showvalidatedonly;
        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortByBtn;
        [FindsBy(How = How.XPath, Using = ALL_FILTER)]
        private IWebElement _all;
        [FindsBy(How = How.XPath, Using = OPENED_FILTER)]
        private IWebElement _opened;
        [FindsBy(How = How.XPath, Using = CLOSED_FILTER)]
        private IWebElement _closed;
        [FindsBy(How = How.Id, Using = Sites_FILTER)]
        private IWebElement _SitesFilter;
        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;
        [FindsBy(How = How.XPath, Using = SITES_SELECT_ALL)]
        private IWebElement _siteSelectAll;
        [FindsBy(How = How.XPath, Using = SITE_UNSELECT_ALL)]
        private IWebElement _siteUnselectAll;
        [FindsBy(How = How.XPath, Using = SITE)]
        private IWebElement _site;
        [FindsBy(How = How.Id, Using = VALIDATE_RESULTS)]
        private IWebElement _validateResults;
        [FindsBy(How = How.Id, Using = CONFIRM_VALIDATE_RESULTS)]
        private IWebElement _confirmValidate;
        [FindsBy(How = How.XPath, Using = NEW_Interim)]
        private IWebElement _createNewInterim;
        [FindsBy(How = How.XPath, Using = CLICKSHOWMENU)]
        private IWebElement _clickshowmenu;
        [FindsBy(How = How.XPath, Using = SUPPLIERS)]
        private IWebElement _suppliers;

        // Tableau
        [FindsBy(How = How.XPath, Using = FIRST_INTERIM)]
        private IWebElement _firstInterim;

        [FindsBy(How = How.XPath, Using = DELETE_INTERIM_BUTTON_DEV)]
        private IWebElement _deleteInterimDev;

        [FindsBy(How = How.XPath, Using = DELETE_INTERIM_CONFIRM_BUTTON)]
        private IWebElement _deleteInterimConfirm;
        [FindsBy(How = How.XPath, Using = EXPORT_BUTTON)]
        private IWebElement _exportbutton;
        [FindsBy(How = How.XPath, Using = EXPORT)]
        private IWebElement _exportbtn;
        [FindsBy(How = How.XPath, Using = FIRST_INTERIM_NUMBER)]
        private IWebElement _firstInterimNumber;

        //filters

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_SEARCH_NUMBER)]
        private IWebElement _searchFilterNumber;

        [FindsBy(How = How.Id, Using = FILTER_SEARC_NUMBER_SECONDARY)]
        private IWebElement _searchFilterNumSecondary;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_SUPPLIER)]
        private IWebElement _supplierFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SUPPLIER)]
        private IWebElement _uncheckAllSupplier;

        [FindsBy(How = How.XPath, Using = SUPPLIER_SEARCH)]
        private IWebElement _searchSupplier;

        [FindsBy(How = How.XPath, Using = SHOW_ALL)]
        private IWebElement _showAll;

        [FindsBy(How = How.XPath, Using = SHOW_ONLY_ACTIVE)]
        private IWebElement _showOnlyActive;

        [FindsBy(How = How.XPath, Using = SHOW_ONLY_INACTIVE)]
        private IWebElement _showOnlyInactive;

        [FindsBy(How = How.XPath, Using = SHOW_ALL_RECEPTIONS)]
        private IWebElement _showAllReceptions;

        [FindsBy(How = How.XPath, Using = SHOW_NOT_VALIDATED_ONLY)]
        private IWebElement _showOnlyNotValidated;

        [FindsBy(How = How.XPath, Using = SHOW_VALIDATED_ONLY)]
        private IWebElement _showOnlyValidated;

        [FindsBy(How = How.XPath, Using = FILTER_INTERIM_RECEPTION_STATUS_ALL)]
        private IWebElement _interimReceptionsAll;

        [FindsBy(How = How.XPath, Using = FILTER_INTERIM_RECEPTION_STATUS_OPENED)]
        private IWebElement _interimReceptionsOpened;

        [FindsBy(How = How.XPath, Using = FILTER_INTERIM_RECEPTION_STATUS_CLOSED)]
        private IWebElement _interimReceptionsClosed;

        [FindsBy(How = How.Id, Using = FILTER_SITES)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_SITES)]
        private IWebElement _searchSites;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SITES)]
        private IWebElement _uncheckAllSites;

        [FindsBy(How = How.Id, Using = FILTER_SORTBY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.XPath, Using = DETAIL_BUTTON)]
        private IWebElement _detail_Button;


        [FindsBy(How = How.XPath, Using = INTERIM_RECEPTION_COUNTER)]
        private IWebElement _interimRecptionCounter;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE)]
        private IWebElement _errormsg;
        [FindsBy(How = How.XPath, Using = SUPPLIER)]
        private IWebElement _supplier;
      
        public enum FilterInterimReceptionStatus
        {
            ALL,
            OPENED,
            CLOSED
        }
        public void ValidateInterimReception()
        {
            Actions action = new Actions(_webDriver);


            var okMenu = WaitForElementIsVisible(By.XPath("//*/button[contains(@class,'dropbtn fas fa-check')]"));
            action.MoveToElement(okMenu).Perform();

            // le menu s'affiche
            WaitPageLoading();
            WaitForLoad();

            var validate = WaitForElementIsVisible(By.Id("btn-validate-interim-reception"));
            validate.Click();
            var validate2 = WaitForElementIsVisible(By.Id("btn-popup-validate"));
            validate2.Click();
        }

        public void WaitForLoad()
        {
            base.WaitForLoad();
        }
        public List<string> GetNameProvidersList()
        {
            var NameProvidersListId = new List<string>();

            var NameProvidersList = _webDriver.FindElements(By.XPath(PROVIDERS_NAMES));

            foreach (var fileFlowProviders in NameProvidersList)
            {
                NameProvidersListId.Add(fileFlowProviders.Text);
            }

            return NameProvidersListId;
        }


        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
        }
        public enum FilterType
        {
            Bynumber,
            Byinterimordernumber,
            Suppliers,
            Showallreceptions,
            Shownotvalidatedonly,
            Showvalidatedonly,
            ShowAll,
            ShowActive,
            ShowInactive,
            SortBy,
            All,
            Opened,
            Closed,
            Sites,
            ByNumber,
            ByInterimOrderNumber,
            DateFrom,
            DateTo,
            ShowOnlyActive,
            ShowOnlyInactive,
            ShowAllReceptions,
            ShowNotValidatedOnly,
            ShowValidatedOnly,
            ShowAllInterimReception,
            ShowOpenedInterimReception,
            ShowClosedInterimReception,
            Site,
            StartDate,
            EndDate,
            SubGroupe


        }
        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.Bynumber:
                    _BynumberFilter = WaitForElementIsVisible(By.Id(BYNUMBERFilter_FILTER));
                    _BynumberFilter.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.Byinterimordernumber:
                    _ByinterimordernumberFilter = WaitForElementIsVisible(By.Id(Byinterimordernumber_FILTER));
                    _ByinterimordernumberFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.Showallreceptions:
                    _showallreceptions = WaitForElementExists(By.XPath(SHOW_ALL_RECEPTIONS_FILTER));
                    _showallreceptions.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.Shownotvalidatedonly:
                    _shownotvalidatedonly = WaitForElementExists(By.XPath(SHOW_VALIDATE_ONLY_FILTER));
                    _shownotvalidatedonly.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.Showvalidatedonly:
                    _showvalidatedonly = WaitForElementExists(By.XPath(SHOW_NOT_VALIDATE_ONLY_FILTER));
                    _showvalidatedonly.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowActive:
                    _showactive = WaitForElementExists(By.XPath(SHOW_ACTIVE_FILTER));
                    _showactive.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowInactive:
                    _showactive = WaitForElementExists(By.XPath(SHOW_ACTIVE_FILTER));
                    _showactive.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.SortBy:
                    _sortByBtn = WaitForElementIsVisible(By.Id(SORTBY_FILTER));
                    _sortByBtn.Click();
                    var elementt = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortByBtn.SetValue(ControlType.DropDownList, elementt.Text);
                    _sortByBtn.Click();
                    break;
                case FilterType.All:
                    _all = WaitForElementExists(By.XPath(ALL_FILTER));
                    _all.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Opened:
                    _opened = WaitForElementExists(By.XPath(OPENED_FILTER));
                    _opened.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Closed:
                    _closed = WaitForElementExists(By.XPath(CLOSED_FILTER));
                    _closed.SetValue(ControlType.RadioButton, value);
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
                case FilterType.ByNumber:
                    _searchFilterNumber = WaitForElementIsVisible(By.Id(FILTER_SEARCH_NUMBER));
                    _searchFilterNumber.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.ByInterimOrderNumber:
                    _searchFilterNumSecondary = WaitForElementIsVisible(By.Id(FILTER_SEARC_NUMBER_SECONDARY));
                    _searchFilterNumSecondary.SetValue(ControlType.TextBox, value);
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

                case FilterType.ShowAllReceptions:
                    _showAllReceptions = WaitForElementExists(By.XPath(SHOW_ALL_RECEPTIONS));
                    action.MoveToElement(_showAllReceptions).Perform();
                    _showAllReceptions.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;

                case FilterType.ShowNotValidatedOnly:
                    _showOnlyNotValidated = WaitForElementExists(By.XPath(SHOW_NOT_VALIDATED_ONLY));
                    action.MoveToElement(_showOnlyNotValidated).Perform();
                    _showOnlyNotValidated.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;

                case FilterType.ShowValidatedOnly:
                    _showOnlyValidated = WaitForElementExists(By.XPath(SHOW_VALIDATED_ONLY));
                    action.MoveToElement(_showOnlyValidated).Perform();
                    _showOnlyValidated.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
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

                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITES, (string)value));
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

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public bool IsSortedBySites()
        {
            var listeSite = _webDriver.FindElements(By.XPath(SITE));

            if (listeSite.Count == 0)
                return false;

            var ancientSite = listeSite[0].Text;

            foreach (var elm in listeSite)
            {
                try
                {
                    if (elm.Text.CompareTo(ancientSite) < 0)
                        return false;

                    ancientSite = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public InterimReceptionsCreateModalPage InterimCreatePage()
        {
            ShowPlusMenu();

            _createNewInterim = WaitForElementIsVisible(By.LinkText(NEW_Interim));
            _createNewInterim.Click();
            WaitForLoad();

            return new InterimReceptionsCreateModalPage(_webDriver, _testContext);
        }
        public InterimReceptionsCreateModalPage CreateNewInterim()
        {
            ShowPlusMenu();

            _createNewInterim = WaitForElementIsVisible(By.XPath(NEW_INTERIM_BUTTON), nameof(NEW_INTERIM_BUTTON));
            _createNewInterim.Click();
            WaitForLoad();

            return new InterimReceptionsCreateModalPage(_webDriver, _testContext);
        }

        public string GetFirstID()
        {
            WaitForLoad();
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
        public void DeleteInterim(string interimNumber)
        {
            _deleteInterimDev = WaitForElementIsVisible(By.XPath(string.Format(DELETE_INTERIM_BUTTON_DEV, interimNumber)), nameof(DELETE_INTERIM_BUTTON_DEV));
            _deleteInterimDev.Click();
            WaitForLoad();

            _deleteInterimConfirm = WaitForElementIsVisible(By.Id(DELETE_INTERIM_CONFIRM_BUTTON));
            _deleteInterimConfirm.Click();
            WaitForLoad();
        }
        public bool CheckValidator()
        {
            return isElementExists(By.XPath(LOCATION_VALIDATOR_MESSAGE)) && isElementExists(By.XPath(DELIVERY_VALIDATOR_MESSAGE)) && isElementExists(By.XPath(SUPPLIER_VALIDATOR_MESSAGE));
        }
        public override void ShowExtendedMenu()
        {
            _clickshowmenu = WaitForElementIsVisible(By.XPath(CLICKSHOWMENU));
            _clickshowmenu.Click();
            WaitForLoad();
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
        public bool IsExcelFileCorrect(string filePath)
        {
            Regex r = new Regex("[LPCarts\\s\\d._]", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public enum ExportType
        {
            Export,
        }
        public void ExportInterimReceptions(ExportType exportType, bool versionPrint)
        {
            ShowExtendedMenu();

            switch (exportType)
            {
                case ExportType.Export:
                    _exportbtn = WaitForElementIsVisible(By.Id(EXPORT));
                    _exportbtn.Click();
                    WaitForLoad();
                    break;


                default:
                    break;
            }

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClosePrintButton();
            }
            WaitForDownload();
            Close();
        }

        public string GetFirstInterimReceptionsNumber()
        {
            _firstInterimNumber = WaitForElementIsVisible(By.XPath(FIRST_INTERIM_NUMBER));
            return _firstInterimNumber.Text;
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
        public object GetFilterValue(FilterType filterType)
        {
            switch (filterType)
            {
                case FilterType.ByNumber:
                    _searchFilterNumber = WaitForElementIsVisible(By.Id(FILTER_SEARCH_NUMBER));
                    return _searchFilterNumber.GetAttribute("value");

                case FilterType.ByInterimOrderNumber:
                    _searchFilterNumSecondary = WaitForElementIsVisible(By.Id(FILTER_SEARC_NUMBER_SECONDARY));
                    return _searchFilterNumSecondary.GetAttribute("value");

                case FilterType.ShowAll:
                    var _showAll = WaitForElementExists(By.XPath(SHOW_ALL));
                    return _showAll.Selected;

                case FilterType.ShowOnlyActive:
                    var _showOnlyActive = WaitForElementIsVisible(By.XPath(SHOW_ONLY_ACTIVE));
                    return _showOnlyActive.Selected;

                case FilterType.ShowOnlyInactive:
                    var _showOnlyInactive = WaitForElementIsVisible(By.XPath(SHOW_ONLY_INACTIVE));
                    return _showOnlyInactive.Selected;

                case FilterType.ShowAllReceptions:
                    var _showAllReceptions = WaitForElementIsVisible(By.XPath(SHOW_ALL_RECEPTIONS));
                    return _showAllReceptions.Selected;

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
        public bool IsSortedByNumber()
        {
            var listeNumber = _webDriver.FindElements(By.Id(BYNUMBERFilter_FILTER));

            if (listeNumber.Count == 0)
                return false;

            var ancientNumber = listeNumber[0].Text;

            foreach (var elm in listeNumber)
            {
                try
                {
                    if (elm.Text.CompareTo(ancientNumber) < 0)
                        return false;

                    ancientNumber = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }
        public bool IsSortedByDeliveryDate(string dateFormat)
        {

            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            ReadOnlyCollection<IWebElement> dates;
            dates = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[5]"));

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
         public bool IsFromToDateRespected(DateTime fromDate, DateTime toDate, string dateFormat)
        {
            var cultureInfo = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var dates = _webDriver.FindElements(By.XPath(DATES));

            if (dates.Count == 0)
                return false;

            return dates
                .Select(elm => DateTime.Parse(elm.Text, cultureInfo).Date)
                .Any(date => date >= fromDate && date <= toDate);
        }

        public List<string> GetAllInterinReceptionPaged()
        {
            var interimRecepElements = _webDriver.FindElements(By.XPath(INTERIM_RECEPTION_TABLE));
            var interimList = interimRecepElements.Select(element => element.Text).ToList();
            return interimList;
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
        public bool IsSortedBySuppliers()
        {
            var listeSuppliers = _webDriver.FindElements(By.XPath(SUPPLIER));

            if (listeSuppliers.Count == 0)
                return false;

            var ancientSuppliers = listeSuppliers[0].Text;

            foreach (var elm in listeSuppliers)
            {
                try
                {
                    if (elm.Text.CompareTo(ancientSuppliers) < 0)
                        return false;

                    ancientSuppliers = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public InterimReceptionsItem ClickFirstLine()
        {
            _detail_Button = WaitForElementIsVisible(By.XPath(DETAIL_BUTTON), nameof(DETAIL_BUTTON));
            _detail_Button.Click();
            WaitForLoad();
            return new InterimReceptionsItem(_webDriver, _testContext);

        }
        public string InterimRecptionCounter()
        {
            _interimRecptionCounter = WaitForElementIsVisible(By.XPath(INTERIM_RECEPTION_COUNTER));
            return _interimRecptionCounter.Text;
        }
        public InterimReceptionsItem GoToInterimReceptionItem()
        {
            ClickFirstInterim();
            WaitForLoad();
            WaitPageLoading();

            return new InterimReceptionsItem(_webDriver, _testContext);
        }
        public void ClickFirstInterim()
        {
            WaitPageLoading();
            _firstInterim = WaitForElementIsVisible(By.XPath(FIRST_INTERIM));
            _firstInterim.Click();
            WaitForLoad();
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
                    _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[1]/div/i[contains(@class,'fa-regular fa-circle-check')]")));

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
        public string GetShowFilterSelected()
        {

            var itemSelected = _webDriver.FindElements(By.XPath(SHOW_SELECTED));
            var showName = string.Empty;
            foreach (var element in itemSelected)
            {
                if (element.Selected)
                {
                    switch (element.GetAttribute("value"))
                    {
                        case "All":
                            showName = "Show all receptions";
                            break;
                        case "NotValidatedOnly":
                            showName = "Show not validated only";
                            break;
                        case "ValidatedOnly":
                            showName = "Show validated only";
                            break;
                        default:
                            showName = "";
                            break;
                    }
                }
            }

            return showName;
        }
        public List<string> GetSelectedSuppliersToFilter()
        {
            var suppliersSelected = new List<string>();
            _suppliers = WaitForElementIsVisible(By.Id(SUPPLIERS));
            _suppliers.Click();
            WaitForLoad();
            var selectElement = _webDriver.FindElements(By.XPath(SUPPLIERS_SELECTED));
            for (var i = 0; i < selectElement.Count; i++)
            {
                if (selectElement[i].Selected)
                {
                    var element = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[10]/ul/li[{0}]/label/input/../span", i + 1)));
                    suppliersSelected.Add(element.Text);
                }
            }

            _suppliers.Click();
            return suppliersSelected;
        }
        public string GetShowFilterActiveSelected()
        {

            var itemSelected = _webDriver.FindElements(By.XPath(SHOW_ACTIVE_SELECTED));
            var showName = string.Empty;
            foreach (var element in itemSelected)
            {
                if (element.Selected)
                {
                    switch (element.GetAttribute("value"))
                    {
                        case "All":
                            showName = "Show all";
                            break;
                        case "ActiveOnly":
                            showName = "Show active only";
                            break;
                        case "InactiveOnly":
                            showName = "Show inactive only";
                            break;
                        default:
                            showName = "";
                            break;
                    }
                }
            }

            return showName;
        }
        public List<string> GetSelectedSitesToFilter()
        {
            var siteSelected = new List<string>();
            _SitesFilter = WaitForElementIsVisible(By.Id(Sites_FILTER));
            _SitesFilter.Click();
            WaitForLoad();
            var selectElement = _webDriver.FindElements(By.XPath(SITES_SELECTED));
            for (var i = 0; i < selectElement.Count; i++)
            {
                if (selectElement[i].Selected)
                {
                    var element = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[11]/ul/li[{0}]/label/input/../span", i + 1)));
                    siteSelected.Add(element.Text);
                }
            }

            _SitesFilter.Click();
            return siteSelected;
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
                        default:
                            showName = "";
                            break;
                    }
                }
            }

            return showName;
        }
        public void ScrollUntilSitesFilterIsVisible()
        {
            ScrollUntilElementIsInView(By.Id(Sites_FILTER));
        }
        public string GetExportErrorMsg()
        {
            _errormsg = WaitForElementIsVisible(By.XPath(ERROR_MESSAGE));
            return _errormsg.Text;
        }

        public void Export()
        {
            ShowExtendedMenu();
            _exportbtn = WaitForElementIsVisible(By.Id(EXPORT));
            _exportbtn.Click();
            WaitForLoad();
        }


        public void ScrollUntilResetFilterIsVisible()
        {
            var element = _webDriver.FindElement(By.XPath(RESET_FILTER));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView({block: 'center'});", element);
            
        }
        public PrintReportPage PrintReport(bool newVersionPrint, bool withPrice=false)
        {
            WaitForLoad();
            ShowExtendedMenu();

            _printinterimreception = WaitForElementIsVisible(By.XPath(PRINTINTERIMRECEPTION));
            _printinterimreception.Click();
            WaitForLoad();

            if (withPrice)
            {
                _ispriceinclude = WaitForElementExists(By.XPath(ISPRICEINCLUDE));                
                if (!_ispriceinclude.Selected)
                {
                    _ispriceinclude.Click();
                }
            }  

            _print_btn = WaitForElementIsVisible(By.XPath(PRINT_BTN));
            _print_btn.Click();
            WaitForLoad();

            if (newVersionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            return new PrintReportPage(_webDriver, _testContext);
        }
        public InterimReceptionsItem ClickFirstlineinterimNotValidated()
        {
            var firstLineElement = _webDriver.FindElement(By.XPath(FIRST_LINE_NOT_VALIDATE));
            firstLineElement.Click();
            WaitForLoad();
            return new InterimReceptionsItem(_webDriver, _testContext);

        }   
        public InterimReceptionsItem SelectInterimReceptionsItem(int index)
        {
            _detail_Button = WaitForElementIsVisible(By.XPath($"//*[@id=\"list-item-with-action\"]/table/tbody/tr[{index}]/td[4]"));
            _detail_Button.Click();
            WaitForLoad();
            return new InterimReceptionsItem(_webDriver, _testContext);

        }
        public bool isPageSizeEqualsTo(int pagesize)
        {
            var nbPages = WaitForElementExists(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div/div[7]/div/div/div/nav/select"));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");
            return selectedValue == pagesize.ToString();
        }

        public int GetNumberOfPageResults()
        {
            if (isElementVisible(By.XPath(string.Format("(//*[@id='list-item-with-action']/nav/ul[@class='pagination']/li/a[not(contains(text(),'>')) and not(contains(text(),'<')) and not(contains(text(),'...'))])[last()]"))))
            {
                var pageNumber = _webDriver.FindElement(By.XPath("(//*[@id='list-item-with-action']/nav/ul[@class='pagination']/li/a[not(contains(text(),'>')) and not(contains(text(),'<')) and not(contains(text(),'...'))])[last()]"));

                return int.Parse(pageNumber.Text);
            }
            return 0;   
        }
        public void GetPageResults(int pagenumber)
        {
            var page = WaitForElementIsVisible(By.XPath("//*[@id='list-item-with-action']/nav/ul/li/a[text()='" + pagenumber + "']"));
            page.Click();
            WaitForLoad();
        }
        public int GetNumberOfShowedResults()
        {
            if (isElementVisible(By.XPath(ALL_RN_ROWS)))
            {
                var pageNumber = _webDriver.FindElements(By.XPath(ALL_RN_ROWS));

                return pageNumber.Count;
            }
            return 0;
        }
        public bool IsChangementPage(int pagenumber)
        {
            if (isElementVisible(By.XPath("//*[@id='list-item-with-action']/nav/ul/li/a[text()='" + pagenumber + "']")))
            {                
                return true;
            }
            return false;
        }

        public InterimReceptionsItem clickResult()
        {
            var choisir = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[2]"));
            choisir.Click();
            return new InterimReceptionsItem(_webDriver, _testContext);
        }

        public IWebElement GetAddNewInterimReceptionButton()
        {
            return WaitForElementIsVisible(By.XPath(ADD_NEW_INTERIM_RECEPTION_BTN));
        }

        public IWebElement GetExtractedNumberElement()
        {
            return WaitForElementIsVisible(By.XPath(EXTRACTED_NUMBER_ELEMENT));
        }

        public IWebElement GetFilterInput()
        {
            return WaitForElementIsVisible(By.XPath(FILTER_INPUT));
        }

        public IList<IWebElement> GetFilteredRows()
        {
            return _webDriver.FindElements(By.XPath(FILTERED_ROWS));
        }

        public IWebElement GetFirstRowCell()
        {
            return WaitForElementIsVisible(By.XPath(FIRST_ROW_CELL));
        }

        public void CreateWithNewQuantity(string site, string supplierName)
        {
            var create2 = WaitForElementIsVisible(By.XPath("//*/button[text()='+']"));
            new Actions(_webDriver).MoveToElement(create2).Perform();
            var createNew = WaitForElementIsVisible(By.XPath("//*/a[text()='New interim reception']"));
            createNew.Click();
            WaitForLoad();

            // checkBoxCopyFrom
            var uncheck = WaitForElementExists(By.Id("checkBoxCopyFrom"));
            uncheck.SetValue(PageBase.ControlType.CheckBox, false);
            var selectSite = WaitForElementIsVisible(By.Id("SelectedSiteId"));
            selectSite.SetValue(PageBase.ControlType.DropDownList, site);
            // select cascade
            WaitForLoad();
            var selectSupplier = WaitForElementIsVisible(By.Id("drop-down-suppliers"));
            selectSupplier.SetValue(PageBase.ControlType.DropDownList, supplierName);
            var selectPlace = WaitForElementIsVisible(By.Id("SelectedSitePlaceId"));
            selectPlace.SetValue(PageBase.ControlType.DropDownList, "Produccion");

            var deliveryOrderNumber = WaitForElementIsVisible(By.Id("InterimReception_DeliveryOrderNumber"));
            Random rnd = new Random();
            deliveryOrderNumber.SetValue(PageBase.ControlType.TextBox, rnd.Next().ToString());

            var create3 = WaitForElementIsVisible(By.XPath("//*/button[text()='Create']"));
            create3.Click();
        }

        public string CheckNumberReception(string decimalSeparator)
        {
            var selectLine = WaitForElementIsVisible(By.XPath("//*/form[@id='itemForm_0']"));
            selectLine.Click();

            var received = WaitForElementIsVisible(By.XPath("//*/form[@id='itemForm_0']/div[1]/div[8]/input"));
            received.Clear();
            received.SendKeys("5" + decimalSeparator + "000");

            ValidateInterimReception();
            var interimReceptionTitle = WaitForElementIsVisible(By.XPath("//*[@id='div-body']/div/div[1]/h1"));
            string interimReceptionNumber = interimReceptionTitle.Text.Substring("INTERIM RECEPTION NO°".Length);

            return interimReceptionNumber;
        }
        public string GetFirstInterimReceptionsDate()
        {
            var firstInterimDate = WaitForElementIsVisible(By.XPath(FIRST_INTERIM_DATE));
            return firstInterimDate.Text;
        }
        public int CheckTotalNumberInterim()
        {
            WaitForLoad();
            var totalNumber = WaitForElementExists(By.ClassName("counter"));
            int nombre = Int32.Parse(totalNumber.GetAttribute("innerText"));
            return nombre;
        }
    }
}
