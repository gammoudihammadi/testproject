using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart
{
    public class LpCartPage : PageBase
    {

        public LpCartPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________ Constantes _____________________________________

        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string PLUS_BTN = "//*[@id=\"tabContentItemContainer\"]/div/div/div[2]/button";
        private const string LIST_ROUTES = "//*[@id=\"tabContentDetails\"]/div[5]/div[2]/div/div[1]";

        private const string NEW_LP_CART = "//*[@id=\"tabContentItemContainer\"]/div/div/div[2]/div/a";
        private const string TRASH_BTN_BIS = "//*[@id=\"lpCart-table\"]/tbody/tr[{0}]/td[11]/a";
        private const string TRASH_BTN = "//*[@id=\"lpCart-table\"]/tbody/tr[2]/td[11]/a";
        private const string CONFIRM_BTN = "first";
        private const string DELETE_CANCEL_BTN = "last";
        private const string DELETE_ERROR_MESSAGE = "ErrorSpan";

        private const string RESET_FILTER = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string SEARCH_FILTER = "SearchPattern";
        private const string SORT_BY_FILTER = "cbSortBy";
        private const string SITES_FILTER = "SelectedSites_ms";
        private const string VERIFY_FILTER_SITE = "/html/body/div[10]/ul/li[*]/label/span";
        private const string UNSELECT_SITES = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_SITE = "/html/body/div[10]/div/div/label/input";
        private const string AIRPORT_FILTER = "cbAirports";
        private const string START_DATE_FILTER = "date-from-picker";
        private const string END_DATE_FILTER = "date-to-picker";
        private const string FROM_HEURES_FILTER = "//*[@id=\"item-filter-form\"]/div[8]/div[1]/input[2]";
        private const string FROM_MINUTES_FILTER = "//*[@id=\"item-filter-form\"]/div[8]/div[1]/input[3]";
        private const string TO_HEURES_FILTER = "//*[@id=\"item-filter-form\"]/div[8]/div[2]/input[2]";
        private const string TO_MINUTES_FILTER = "//*[@id=\"item-filter-form\"]/div[8]/div[2]/input[3]";
        private const string ROUTE_FILTER = "Route";
        private const string EXPIRED_SERVICES_FILTER = "ShowLPWithExpiredServices";
        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string UNSELECT_CUSTOMERS = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_CUSTOMER = "/html/body/div[11]/div/div/label/input";
        private const string EXTENDED_BTN = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/button";
        private const string DUPLICATE_LPCART = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/div/a[2]";
        private const string EXPORT = "btn-export-excel";
        private const string CHOOSE_FILE = "fileSent";
        private const string VALIDATE_IMPORT = "//*[@id=\"ImportFileForm\"]/div[2]/button[2]";
        private const string OK_IMPORT = "//*[@id=\"modal-1\"]/div/div/div[3]/button";
        private const string IMPORT_LP_CART = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/div/a[1]";
        private const string BY_VALIDITY_DATE = "/html/body/div[2]/div/div[1]/div[2]/div/form/div[6]/input[1]";
        private const string GENERAL_INFO_TAB = "/html/body/div[3]/div/ul/li[1]/a";
        private const string ROUTE_SELECT = "/html/body/div[3]/div/div[2]/div/div/div/form/div/div[5]/div[2]/div/div[1]";
        private const string CARTS_TAB = "/html/body/div[3]/div/ul/li[2]/a";
        private const string NUMBER_LPCART = "/html/body/div[2]/div/div[2]/div/div/h1/span";
        //Cart detail
        private const string FIRST_LPCART_NAME = "//*[@id=\"lpCart-table\"]/tbody/tr[2]/td[2]";
        private const string FIRST_LPCART_CODE = "//*[@id=\"lpCart-table\"]/tbody/tr[2]/td[1]";
        private const string ITEM_NAME = "//*[@id=\"lpCart-table\"]/tbody/tr[*]/td[2]";
        private const string DATE_FROM = "//*[@id=\"lpCart-table\"]/tbody/tr[*]/td[6]";
        private const string VERIFY_SITE = "//*[@id=\"lpCart-table\"]/tbody/tr[*]/td[3]";
        private const string VERIFY_CUSTOMER = "//*[@id=\"lpCart-table\"]/tbody/tr[*]/td[4]";
        private const string FROM_DATE = "//*[@id=\"lpCart-table\"]/tbody/tr[*]/td[6]";
        private const string TO_DATE = "//*[@id=\"lpCart-table\"]/tbody/tr[*]/td[7]";
        private const string ITEM = "//*[@id=\"lpCart-table\"]/tbody/tr[2]";
        private const string FLIGHT_COUNT = "//*[@id=\"lpCart-table\"]/tbody/tr[*]/td[8]";
        private const string ROUTES = "/html/body/div[3]/div/div[2]/div/div/div/form/div/div[5]/div[1]/div/div/span/input[2]";
        private const string DELETE_ROUTE = "/html/body/div[3]/div/div[2]/div/div/div/form/div/div[5]/div[2]/div/div[1]/a/span";
        private const string ADD_ROUTE = "/html/body/div[3]/div/div[2]/div/div/div/form/div/div[5]/div[1]/div/div/a";


        //Lpcart scheme
        private const string PLUS_BTN_LP_CART_SCHEME = "//span[contains(@class, 'glyphicon glyphicon-plus')]";
        private const string PLUS_BTN_LP_CART_SCHEME_DEV = "//span[contains(@title, 'Add a new cart scheme')]/a";
        private const string EDIT_BTN_LP_CART_SCHEME = "//span[contains(@class, 'glyphicon glyphicon-pencil')]";
        private const string EDIT_BTN_LP_CART_SCHEME_DEV = "//span[contains(@title, 'Edit cart scheme')]/a";
        private const string LP_CART_FROM_ALL = "//*[@id=\"lpCart-table\"]/tbody/tr[*]/td[6]";
        private const string FILTER_BY_DATE = "/html/body/div[2]/div/div[1]/div[2]/div/form/div[6]/input[2]";

        private const string LP_CART_TROLLEY_NAMES = "/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[*]/td[3]";
        private const string FILTER_START_DATE = "//*[@id=\"date-from-picker\"]";
        private const string FILTER_END_DATE = "//*[@id=\"date-to-picker\"]";

        private const string FROM_DATES = "//*[@id=\"lpCart-table\"]/tbody/tr[*]/td[6]";
        private const string TO_DATES = "//*[@id=\"lpCart-table\"]/tbody/tr[*]/td[7]";
        private const string LPCART_AIRCRAFTS = "//*[@id=\"selected-aircrafts\"]/div[*]/span";
        private const string CONFIRM_DELETE_LPCART = "//*[@id=\"first\"]";
        private const string DUPLICATE_CONFIRM_DELETE_LPCART = "/html/body/div[2]/div/div[2]/div/table/tbody/tr[2]/td[11]/a";

        private const string GENERATE_DRAWER_BUTTON = "btnGenerate";
        private const string GENERATE_DRAWER_CONFIRM_BUTTON = "dataConfirmOK";
        private const string LP_CART_CODE = "tb-new-lpcart-number";
        private const string LP_CART_NAME = "tb-new-lpcart-name";
        private const string LP_CART_SITE = "//*[@id=\"input-siteCode\"]";
        private const string LP_CART_CUSTOMER = "//*[@id=\"input-customerCodeAndName\"]";
        private const string LP_CART_AIRCRAFT = "//*[@id=\"selected-aircrafts\"]/div/span";
        private const string LP_CART_FROM = "datapicker-edit-lpcart-from";
        private const string LP_CART_TO = "datapicker-edit-lpcart-to";
        private const string LP_CART_COMMENT = "Comment";
        private const string AIRCRAFT_TAGS = "//*[@id=\"selected-aircrafts\"]/div[*]";
        private const string ROUTES_TAGS = "//*[@id=\"tabContentDetails\"]/div[5]/div[2]/div/div[*]";

        private const string SCHEMA_LINE_COUNT = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[2]/td[9]";
        private const string SCHEMA_COLUMN_COUNT = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[2]/td[10]";
        private const string LP_CART_EDIT_TITLE = "//*[@id=\"Title\"]";
        private const string LP_CART_EDIT_PREPZONE = "//*[@id=\"PrepZoneConcat\"]";
        private const string LP_CART_EDIT_EQPMNAMELABEL = "//*[@id=\"EquipmentNameLabel\"]";
        private const string LP_CART_EDIT_COMMENT = "//*[@id=\"Comment\"]";
        private const string LP_CART_EDIT_SALS_NUMBER = "//*[@id=\"sealsNumber\"]";
        private const string LP_CART_EDIT_LABEL_PAGE = "//*[@id=\"labelPageNbr\"]";
        private const string LP_CART_EDIT_CART_COLOR = "//*[@id=\"trolley-label-color\"]";
        private const string LP_CART_EDIT_MAIN_POSITION = "//*[@id=\"GlobalPosition\"]";
        private const string LP_CART_EDIT_ROWS = "//*[@id=\"rowsNumber\"]";
        private const string LP_CART_EDIT_COLUMNS = "//*[@id=\"colsNumber\"]";
        private const string LP_CART_AIRCRAFTS_TYPE = "//*[@id=\"lpCart-table\"]/tbody/tr[2]/td[5]/p";
        private const string SITE_LP_CART_GENERAL_INFOR = "input-siteCode";
        private const string CUSTOMER_LP_CART_GENERAL_INFOR = "input-customerCodeAndName";
        private const string LP_CART_NAMES = "//*[@id=\"lpCart-table\"]/tbody/tr[*]/td[2]";
        private const string FLIGHTS_COUNT = "//*[@id=\"lpCart-table\"]/tbody/tr[2]/td[8]";

        // __________________________________ Variables ______________________________________

        [FindsBy(How = How.XPath, Using = LP_CART_EDIT_TITLE)]
        private IWebElement _titleEditLpCart;

        [FindsBy(How = How.XPath, Using = LP_CART_EDIT_PREPZONE)]
        private IWebElement _prepzoneEditLpCart;

        [FindsBy(How = How.XPath, Using = LP_CART_EDIT_EQPMNAMELABEL)]
        private IWebElement _eqpmnamelabelEditLpCart;

        [FindsBy(How = How.XPath, Using = LP_CART_EDIT_COMMENT)]
        private IWebElement _commentEditLpCart;

        [FindsBy(How = How.XPath, Using = LP_CART_EDIT_SALS_NUMBER)]
        private IWebElement _numberEditLpCart;

        [FindsBy(How = How.XPath, Using = LP_CART_EDIT_LABEL_PAGE)]
        private IWebElement _labelPageNbrEditCart;

        [FindsBy(How = How.XPath, Using = LP_CART_EDIT_CART_COLOR)]
        private IWebElement _cartColorEditCart;

        [FindsBy(How = How.XPath, Using = LP_CART_EDIT_MAIN_POSITION)]
        private IWebElement _mainPositionEditCart;

        [FindsBy(How = How.XPath, Using = LP_CART_EDIT_ROWS)]
        private IWebElement _rowsEditCart;

        [FindsBy(How = How.XPath, Using = LP_CART_EDIT_COLUMNS)]
        private IWebElement _columnEditCart;

        [FindsBy(How = How.XPath, Using = NEW_LP_CART)]
        private IWebElement _createNewLpCart;

        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _showPlusNewLpCart;

        [FindsBy(How = How.XPath, Using = DUPLICATE_LPCART)]
        private IWebElement _duplicateLpCart;

        [FindsBy(How = How.XPath, Using = TRASH_BTN)]
        private IWebElement _btn_Trash;

        [FindsBy(How = How.Id, Using = CONFIRM_BTN)]
        private IWebElement _confirm;

        [FindsBy(How = How.Id, Using = DELETE_CANCEL_BTN)]
        private IWebElement _cancel;

        [FindsBy(How = How.XPath, Using = FIRST_LPCART_NAME)]
        private IWebElement _firstLpCartName;

        [FindsBy(How = How.XPath, Using = FIRST_LPCART_CODE)]
        private IWebElement _firstLpCartCode;

        //Filter
        [FindsBy(How = How.Id, Using = SITES_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.Id, Using = UNSELECT_SITES)]
        private IWebElement _unselectAllSites;

        [FindsBy(How = How.Id, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.XPath, Using = ITEM)]
        private IWebElement _firstItem;

        [FindsBy(How = How.XPath, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.XPath, Using = IMPORT_LP_CART)]
        private IWebElement _importLpCart;

        [FindsBy(How = How.Id, Using = CHOOSE_FILE)]
        private IWebElement _chooseFile;

        [FindsBy(How = How.XPath, Using = VALIDATE_IMPORT)]
        private IWebElement _validateImport;

        [FindsBy(How = How.XPath, Using = OK_IMPORT)]
        private IWebElement _okImport;

        [FindsBy(How = How.XPath, Using = PLUS_BTN_LP_CART_SCHEME)]
        private IWebElement _addNewLpCartScheme;

        [FindsBy(How = How.XPath, Using = EDIT_BTN_LP_CART_SCHEME)]
        private IWebElement _editLpCartScheme;

        [FindsBy(How = How.XPath, Using = SCHEMA_COLUMN_COUNT)]
        private IWebElement _schemeColumnCount;

        [FindsBy(How = How.XPath, Using = SCHEMA_LINE_COUNT)]
        private IWebElement _schemeLineCount;
        [FindsBy(How = How.XPath, Using = FLIGHTS_COUNT)]
        private IWebElement _flightsCount;

        // __________________________________ Filtres ________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _search;

        [FindsBy(How = How.Id, Using = SORT_BY_FILTER)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = START_DATE_FILTER)]
        private IWebElement _startDate;

        [FindsBy(How = How.Id, Using = END_DATE_FILTER)]
        private IWebElement _endDate;

        [FindsBy(How = How.Id, Using = EXPIRED_SERVICES_FILTER)]
        private IWebElement _expiredServiceInput;

        [FindsBy(How = How.Id, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_CUSTOMERS)]
        private IWebElement _unselectAllCustomers;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.XPath, Using = EXTENDED_BTN)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.XPath, Using = GENERATE_DRAWER_BUTTON)]
        private IWebElement _generateButton;

        [FindsBy(How = How.XPath, Using = GENERATE_DRAWER_CONFIRM_BUTTON)]
        private IWebElement _validateGeneration;

        public enum FilterType
        {
            Search,
            SortBy,
            Site,
            Customers,
            StartDate,
            EndDate,
            ValidityDate,
            ByFlightdate

        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.Search:
                    _search = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _search.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORT_BY_FILTER));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;
                case FilterType.Site:
                    _siteFilter = WaitForElementExists(By.Id(SITES_FILTER));
                    action.MoveToElement(_siteFilter).Perform();
                    _siteFilter.Click();

                    // On décoche toutes les options
                    _unselectAllSites = _webDriver.FindElement(By.XPath(UNSELECT_SITES));
                    _unselectAllSites.Click();

                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                    _searchSite.SetValue(ControlType.TextBox, value);

                    var valueToCheckSite = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    valueToCheckSite.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();

                    _siteFilter.Click();
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

                case FilterType.ValidityDate:
                    _expiredServiceInput = WaitForElementExists(By.XPath(BY_VALIDITY_DATE));
                    action.MoveToElement(_expiredServiceInput).Perform();
                    _expiredServiceInput.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.ByFlightdate:
                    _expiredServiceInput = WaitForElementExists(By.XPath(FILTER_BY_DATE));
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


        public enum ExportType
        {
            Export,
        }

        public void ExportLpcart(ExportType exportType, bool versionPrint)
        {
            ShowExtendedMenu();

            switch (exportType)
            {
                case ExportType.Export:
                    _export = WaitForElementIsVisible(By.Id(EXPORT));
                    _export.Click();
                    // line auto refresh ?
                    Thread.Sleep(2000);
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
        }

        //___________________________________ Méthodes ___________________________________________

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

        public void Import(string filePath)
        {
            ShowExtendedMenu();

            _importLpCart = WaitForElementIsVisible(By.XPath(IMPORT_LP_CART));
            _importLpCart.Click();
            // animation : la fenêtre arrive
            //Thread.Sleep(2000);
            WaitForLoad();

            _chooseFile = WaitForElementIsVisible(By.Id(CHOOSE_FILE));
            _chooseFile.SendKeys(filePath);
            // le modal grossi lorsqu'on sélectionne le fichier
            //Thread.Sleep(2000);
            WaitForLoad();

            _validateImport = WaitForElementIsVisible(By.XPath(VALIDATE_IMPORT));
            _validateImport.Click();
            // traitement en cours
            //Thread.Sleep(2000);
            WaitForLoad();

            _okImport = WaitForElementIsVisible(By.XPath("//*[@id='modal-1']/div[3]/button"));
            _okImport.Click();
            // attendre line auto refresh
            Thread.Sleep(2000);
            WaitForLoad();

        }

        public void DeleteLpCart()
        {
            if (CheckTotalNumber() > 0)
            {
                if (!isPageSizeEqualsTo100())
                {
                    PageSize("100");
                }
                var lpCartNumber = CheckTotalNumber();

                int i = 2;
                while (i < lpCartNumber + 2)
                {
                    _btn_Trash = WaitForElementIsVisible(By.XPath(String.Format(TRASH_BTN_BIS, i)));
                    _btn_Trash.Click();
                    WaitForLoad();
                    _confirm = WaitForElementIsVisible(By.Id(CONFIRM_BTN));
                    _confirm.Click();
                    WaitForLoad();
                    if (isElementVisible(By.Id(DELETE_ERROR_MESSAGE)))
                    {
                        _cancel = WaitForElementIsVisible(By.Id(DELETE_CANCEL_BTN));
                        _cancel.Click();
                        WaitForLoad();
                        i++;
                    }
                    lpCartNumber = CheckTotalNumber();
                }
            }
        }


        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //date à null
            }
        }

        public string GetFirstLpCartName()
        {
            _firstLpCartName = WaitForElementIsVisible(By.XPath(FIRST_LPCART_NAME));
            return _firstLpCartName.Text;
        }

        public string GetFirstLpCartCode()
        {
            _firstLpCartCode = WaitForElementIsVisible(By.XPath(FIRST_LPCART_CODE));
            return _firstLpCartCode.Text;
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


            foreach (var elem in elements)
            {

                try
                {
                    if (String.Compare(ancienName, elem.Text) > 0)
                    { valueBool = false; }

                    ancienName = elem.Text;
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

        public bool VerifySite(string value)
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

            var elements = _webDriver.FindElements(By.XPath(VERIFY_SITE));


            foreach (var element in elements)
            {
                if (element.Text != value)
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool VerifyFilterSite(string value)
        {
            bool valueBool = true;

            var elements = _webDriver.FindElements(By.XPath(VERIFY_FILTER_SITE));

            foreach (var element in elements)
            {
                if (element.Text.Contains(value))
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

        public bool IsDateRespected(DateTime fromDate, DateTime toDate, string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var datesFrom = _webDriver.FindElements(By.XPath(FROM_DATE));
            var datesTo = _webDriver.FindElements(By.XPath(TO_DATE));

            if (datesFrom.Count == 0)
                return false;

            if (datesTo.Count == 0)
                return false;

            var checkNumber = CheckTotalNumber();
            //si lpcartFrom < From et lpcartTo < From return False
            //si lpcartfrom > To et lpcartTo > To return false

            for (int i = 0; i < checkNumber; i++)
            {

                var datesFromtest = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div/table/tbody/tr[" + (i + 2) + "]/td[6]"));
                var datesTotest = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div/table/tbody/tr[" + (i + 2) + "]/td[7]"));

                DateTime datef = DateTime.Parse(datesFromtest.Text, ci).Date;
                DateTime datet = DateTime.Parse(datesTotest.Text, ci).Date;
                if (DateTime.Compare(datef, fromDate) < 0 && DateTime.Compare(datet, fromDate) < 0 || DateTime.Compare(datef, toDate) > 0 && DateTime.Compare(datet, toDate) > 0)
                {
                    return false;
                }
            }

            return true;
        }

        public LpCartCartDetailPage ClickFirstLpCart()
        {

            var firstLpcart = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div/table/tbody/tr[2]"));
            firstLpcart.Click();
            WaitForLoad();

            return new LpCartCartDetailPage(_webDriver, _testContext);
        }

        public string GetSchemeColumnCount()
        {
            _schemeColumnCount = WaitForElementIsVisible(By.XPath(SCHEMA_COLUMN_COUNT));
            return _schemeColumnCount.Text;
        }
        public string GetSchemaLineCount()
        {
            _schemeLineCount = WaitForElementIsVisible(By.XPath(SCHEMA_LINE_COUNT));
            return _schemeLineCount.Text;
        }

        public string GetTitleEditLpCart()
        {
            _titleEditLpCart = WaitForElementIsVisible(By.XPath(LP_CART_EDIT_TITLE));
            return _titleEditLpCart.GetAttribute("value");
        }

        public string GetPrepzoneEditLpCart()
        {
            _prepzoneEditLpCart = WaitForElementIsVisible(By.XPath(LP_CART_EDIT_PREPZONE));
            return _prepzoneEditLpCart.GetAttribute("value");
        }
        public string GetCommentEditLpCart()
        {
            _commentEditLpCart = WaitForElementIsVisible(By.XPath(LP_CART_EDIT_COMMENT));
            return _commentEditLpCart.GetAttribute("value");
        }

        public string GetEqpNameLabelEditLpCart()
        {
            _eqpmnamelabelEditLpCart = WaitForElementIsVisible(By.XPath(LP_CART_EDIT_EQPMNAMELABEL));
            return _eqpmnamelabelEditLpCart.GetAttribute("value");
        }

        public string GetSalsNumberEditLpCart()
        {
            _numberEditLpCart = WaitForElementIsVisible(By.XPath(LP_CART_EDIT_SALS_NUMBER));
            return _numberEditLpCart.GetAttribute("value");
        }

        public string GetLabelPageNbrEditLpCart()
        {
            _labelPageNbrEditCart = WaitForElementIsVisible(By.XPath(LP_CART_EDIT_LABEL_PAGE));
            return _labelPageNbrEditCart.GetAttribute("value");
        }

        public string GetCartColorEditLpCart()
        {
            _cartColorEditCart = WaitForElementIsVisible(By.XPath(LP_CART_EDIT_CART_COLOR));
            return _cartColorEditCart.GetAttribute("value");
        }

        public string GetMainPositionEditLpCart()
        {
            _mainPositionEditCart = WaitForElementIsVisible(By.XPath(LP_CART_EDIT_MAIN_POSITION));
            return _mainPositionEditCart.GetAttribute("value");
        }

        public string GetRowsNumberEditLpCart()
        {
            _rowsEditCart = WaitForElementIsVisible(By.XPath(LP_CART_EDIT_ROWS));
            return _rowsEditCart.GetAttribute("value");
        }

        public string GetColumnsNumberEditLpCart()
        {
            _columnEditCart = WaitForElementIsVisible(By.XPath(LP_CART_EDIT_COLUMNS));
            return _columnEditCart.GetAttribute("value");
        }
        public string GetAircraftTypeLpCart()
        {
            var _columnAircraftLpCart = WaitForElementIsVisible(By.XPath(LP_CART_AIRCRAFTS_TYPE));
            return _columnAircraftLpCart.Text;
        }

        //___________________________________ Pages ____________________________________________

        public LpCartCreateModalPage LpCartCreatePage()
        {

            _showPlusNewLpCart = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            _showPlusNewLpCart.Click();
            WaitForLoad();

            _createNewLpCart = WaitForElementIsVisible(By.XPath(NEW_LP_CART));
            _createNewLpCart.Click();
            WaitForLoad();

            return new LpCartCreateModalPage(_webDriver, _testContext);
        }

        public LpCartCartDetailPage LpCartCartDetailPage()
        {

            _firstItem = WaitForElementIsVisible(By.XPath(ITEM));
            _firstItem.Click();
            WaitForLoad();

            return new LpCartCartDetailPage(_webDriver, _testContext);
        }

        public LpCartCreateLpCartSchemeModal LpCartCreateLpCartSchemeModal()
        {
            _addNewLpCartScheme = WaitForElementIsVisible(By.XPath(PLUS_BTN_LP_CART_SCHEME_DEV));
            _addNewLpCartScheme.Click();
            WaitForLoad();

            return new LpCartCreateLpCartSchemeModal(_webDriver, _testContext);
        }

        public LpCartEditLpCartSchemeModal LpCartEditLpCartSchemeModal()
        {
            _editLpCartScheme = WaitForElementIsVisible(By.XPath(EDIT_BTN_LP_CART_SCHEME_DEV));
            _editLpCartScheme.Click();
            WaitForLoad();

            return new LpCartEditLpCartSchemeModal(_webDriver, _testContext);
        }

        public void GenerateDrawers()
        {
            _generateButton = WaitForElementIsVisible(By.Id("btnGenerate"));
            _generateButton.Click();
            WaitForLoad();

            _validateGeneration = WaitForElementIsVisible(By.Id("dataConfirmOK"));
            _validateGeneration.Click();
            WaitForLoad();
        }

        public LpCartDuplicateLpCartModalPage DuplicateLpCart()
        {
            ShowExtendedMenu();

            _duplicateLpCart = WaitForElementIsVisible(By.XPath(DUPLICATE_LPCART));
            _duplicateLpCart.Click();
            WaitForLoad();

            return new LpCartDuplicateLpCartModalPage(_webDriver, _testContext);
        }

        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);

            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BTN));
            actions.MoveToElement(_extendedButton).Perform();
        }

        public string GetLPCartDeleteErrorMsg()
        {
            string errorMessage = null;
            if (CheckTotalNumber() > 0)
            {
                if (!isPageSizeEqualsTo100())
                {
                    PageSize("100");
                }

                var lpCartNumber = CheckTotalNumber();

                for (int i = 0; i < lpCartNumber; i++)
                {
                    _btn_Trash = WaitForElementIsVisible(By.XPath(TRASH_BTN));
                    _btn_Trash.Click();
                    WaitForLoad();
                    _confirm = WaitForElementIsVisible(By.Id(CONFIRM_BTN));
                    _confirm.Click();
                    WaitForLoad();
                    if (isElementVisible(By.Id("ErrorSpan")))
                    {
                        errorMessage = WaitForElementIsVisible(By.Id("ErrorSpan")).Text;
                        _confirm = WaitForElementIsVisible(By.Id(DELETE_CANCEL_BTN));
                        _confirm.Click();
                        WaitForLoad();
                        return errorMessage;
                    }
                }
            }
            return errorMessage;
        }

        public bool VerifyLpCartExist(string lpCartName)
        {
            var firstLpCart = isElementExists(By.XPath(FIRST_LPCART_NAME));
            if (firstLpCart)
            {
                var firstLpCartName = WaitForElementIsVisible(By.XPath(FIRST_LPCART_NAME)).Text;
                if (firstLpCartName == lpCartName)
                {
                    return true;
                }
            }
            return false;
        }

        public bool VerifierFilterByDate()
        {
            var allFlightsFromDate = _webDriver.FindElements(By.XPath(LP_CART_FROM_ALL));
            for (int i = 0; i < allFlightsFromDate.Count - 1; i++)
            {

                DateTime date1 = DateTime.ParseExact(allFlightsFromDate[i].Text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
                DateTime date2 = DateTime.ParseExact(allFlightsFromDate[i + 1].Text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
                int result = DateTime.Compare(date1, date2);

                if (result > 0)
                {
                    return false;
                }
            }
            return true;
        }

        public bool verifierImport(string value)
        {


            var trolleyNames = _webDriver.FindElements(By.XPath(LP_CART_TROLLEY_NAMES));
            foreach (var name in trolleyNames)
            {
                if (name.Text != value)
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifyFilterByValidityDate()
        {
            string format = "dd/MM/yyyy";
            var startDate = WaitForElementIsVisible(By.XPath(FILTER_START_DATE)).GetAttribute("value");
            var endDate = WaitForElementIsVisible(By.XPath(FILTER_END_DATE)).GetAttribute("value");
            var startDateTime = DateTime.ParseExact(startDate, format, CultureInfo.InvariantCulture);
            var endDateTime = DateTime.ParseExact(endDate, format, CultureInfo.InvariantCulture);
            var fromDatesInTable = _webDriver.FindElements(By.XPath(FROM_DATES));
            var toDatesInTable = _webDriver.FindElements(By.XPath(TO_DATES));
            for (int i = 0; i < fromDatesInTable.Count; i++)
            {
                // échéances
                //[29/08/2024:29/09/2025]
                //si 29/09/2025 < 01_09_2024
                //si 29/08/2024 > 28_11_2024
                //[01_09_2024:28_11_2024]
                if (((DateTime.ParseExact(toDatesInTable[i].Text, format, CultureInfo.InvariantCulture) < startDateTime) ||
                        (DateTime.ParseExact(fromDatesInTable[i].Text, format, CultureInfo.InvariantCulture) > endDateTime)))
                {
                    return false;
                }
            }
            return true;
        }

        public IEnumerable<string> GetAircraftsByLpCartName(string lpCartName)
        {
            var homePage = new HomePage(_webDriver, _testContext);
            homePage.Navigate();
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.Filter(FilterType.Search, lpCartName);
            var lpCartDetailsPage = ClickFirstLpCart();
            var generalInfoPage = lpCartDetailsPage.LpCartGeneralInformationPage();
            var aircraftsWebElements = _webDriver.FindElements(By.XPath(LPCART_AIRCRAFTS));
            return aircraftsWebElements.Select(e => e.Text).ToList();
        }

        public void RemoveDuplicate(string duplicateName)
        {
            var homePage = new HomePage(_webDriver, _testContext);
            homePage.Navigate();
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.Filter(FilterType.Search, duplicateName);
            lpCartPage.PageSize("100");
            var deleteBtn = WaitForElementIsVisible(By.XPath(DUPLICATE_CONFIRM_DELETE_LPCART));
            deleteBtn.Click();
            WaitForElementIsVisible(By.XPath(CONFIRM_DELETE_LPCART)).Click();

        }



        public bool VerifRouteSelect(string routeInput)
        {
            var generalInformationTab = WaitForElementIsVisible(By.XPath(GENERAL_INFO_TAB));
            generalInformationTab.Click();
            WaitLoading();
            var routeSelect = WaitForElementIsVisible(By.XPath(ROUTE_SELECT));
            var routeVerif = routeSelect.Text.Trim();
            WaitLoading();
            if (routeVerif != routeInput.Trim())
            {
                return false;
            }
            return true;
        }

        public void AddNewRouteLpCartDetail(string newRouteInput)
        {
            var routeDelete = WaitForElementIsVisible(By.XPath(DELETE_ROUTE));
            routeDelete.Click();

            var routesSaisie = WaitForElementIsVisible(By.XPath(ROUTES));
            routesSaisie.SetValue(ControlType.TextBox, newRouteInput);
            var addRoute = WaitForElementIsVisible(By.XPath(ADD_ROUTE));
            addRoute.Click();
            WaitPageLoading();
        }

        public void DeleteCodeLpCartTestIfExist(string code, LpCartPage lpCartPage)
        {
            lpCartPage.Filter(FilterType.Search, code);
            WaitForLoad();

            var numberLpCart = WaitForElementIsVisible(By.XPath(NUMBER_LPCART));
            if (numberLpCart.Text != "0")
            {
                var deleteBtn = WaitForElementIsVisible(By.XPath(DUPLICATE_CONFIRM_DELETE_LPCART));
                deleteBtn.Click();
                WaitForLoad();

                WaitForElementIsVisible(By.XPath(CONFIRM_DELETE_LPCART)).Click();
                WaitPageLoading();

                ResetFilter();
                WaitForLoad();
            }
            ResetFilter();
            WaitForLoad();
        }
        public bool VerifyAfterChangedTab(string route)
        {
            var actions = new Actions(_webDriver);
            var cartsTab = WaitForElementIsVisible(By.XPath(CARTS_TAB));
            actions.MoveToElement(cartsTab).Perform();
            cartsTab.Click();
            WaitForLoad();
            var generalInfo = WaitForElementIsVisible(By.XPath(GENERAL_INFO_TAB));
            generalInfo.Click();
            WaitForLoad();

            var verified = VerifRouteSelect(route);
            WaitForLoad();
            if (verified)
            {
                return true;
            }
            return false;
        }
        public void okWarning()
        {
            if (isElementVisible(By.XPath("//*[@id=\"dataAlertCancel\"]")))
            {
                var okWarning = WaitForElementIsVisible(By.XPath("//*[@id=\"dataAlertCancel\"]"));
                okWarning.Click();
            }
        }

        public bool VerifCreateLPCartInfo(string lpCartCode, string lpCartName, string lpCartSite, string lpCartCustomer, string lpCartAircraft, DateTime lpCartFrom, DateTime lpCartTo, string lpCartComment)
        {
            var generalInformationTab = WaitForElementIsVisible(By.XPath(GENERAL_INFO_TAB));
            generalInformationTab.Click();
            WaitPageLoading();
            var code = WaitForElementIsVisible(By.Id(LP_CART_CODE)).GetAttribute("value");
            var name = WaitForElementIsVisible(By.Id(LP_CART_NAME)).GetAttribute("value");
            var site = WaitForElementIsVisible(By.Id(SITE_LP_CART_GENERAL_INFOR)).GetAttribute("value");
            var customer = WaitForElementIsVisible(By.Id(CUSTOMER_LP_CART_GENERAL_INFOR)).GetAttribute("value");
            var aircraft = WaitForElementExists(By.XPath(LP_CART_AIRCRAFT)).Text;
            var from = DateTime.ParseExact(WaitForElementExists(By.Id(LP_CART_FROM)).GetAttribute("value"),"dd/MM/yyyy",CultureInfo.InvariantCulture);
            var to = DateTime.ParseExact(WaitForElementExists(By.Id(LP_CART_TO)).GetAttribute("value"),"dd/MM/yyyy", CultureInfo.InvariantCulture);
            var comment = WaitForElementExists(By.Id(LP_CART_COMMENT)).GetAttribute("value");
            return IsMatch(lpCartCode,code) && IsMatch(lpCartName, name) && IsMatch(lpCartSite ,site) && IsMatch(lpCartCustomer,customer)
                && IsMatch(lpCartAircraft, aircraft) && DateTime.Equals(lpCartFrom, from) && DateTime.Equals(lpCartTo, to)
                && IsMatch(lpCartComment ,comment);
        }
        public bool VerifCreateLPCartInfoWithRoutes(string lpCartCode, string lpCartName, string lpCartSite, string lpCartCustomer, List<string> lpCartAircraft, DateTime lpCartFrom, DateTime lpCartTo, string lpCartComment, List<string> lpCartRoutes)
        {
            var generalInformationTab = WaitForElementIsVisible(By.XPath(GENERAL_INFO_TAB));
            generalInformationTab.Click();
            WaitPageLoading();

            var code = WaitForElementIsVisible(By.Id(LP_CART_CODE)).GetAttribute("value");
            var name = WaitForElementIsVisible(By.Id(LP_CART_NAME)).GetAttribute("value");
            var site = WaitForElementIsVisible(By.XPath(LP_CART_SITE)).GetAttribute("value");
            var customer = WaitForElementIsVisible(By.XPath(LP_CART_CUSTOMER)).GetAttribute("value");
            var from = DateTime.ParseExact(WaitForElementIsVisible(By.Id(LP_CART_FROM)).GetAttribute("value"), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            var to = DateTime.ParseExact(WaitForElementIsVisible(By.Id(LP_CART_TO)).GetAttribute("value"), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            var comment = WaitForElementExists(By.Id(LP_CART_COMMENT)).GetAttribute("value");
            WaitLoading();
            return IsMatch(lpCartCode, code)
        && IsMatch(lpCartName, name)
        && IsMatch(lpCartSite, site)
        && IsMatch(lpCartCustomer, customer)
        && AreListsMatching(lpCartAircraft, GetAircrafts())
        && IsMatch(lpCartFrom, from)
        && IsMatch(lpCartTo, to)
        && IsMatch(lpCartComment, comment)
        && AreListsMatching(lpCartRoutes, GetRoutes());
        }
        private List<string> GetRoutes()
        {
            var routesElements = _webDriver.FindElements(By.XPath(ROUTES_TAGS));
            return routesElements.Select(element => element.Text).ToList();
        }

        public bool VerifLPCartHasRoutes(List<string> expectedRoutes)
        {
            WaitPageLoading();
            var actualRoutes = GetRoutes();

            // Vérifier que les listes correspondent
            return AreListsMatching(expectedRoutes, actualRoutes);
        }

        public List<string> GetRouteNames()
        {
            var routes = _webDriver.FindElements(By.XPath(LIST_ROUTES));
            return routes.Select(route => route.Text).ToList();
        }

        public void VerifyRoutesNamesUnchanged(List<string> oldRouteNames)
        {
            var newRouteNames = GetRouteNames();
            Assert.IsTrue(oldRouteNames.SequenceEqual(newRouteNames), "Les noms des routes ont changé après modification du nom du LP Cart.");
        }

        private List<string> GetAircrafts()
        {
            return _webDriver.FindElements(By.XPath(AIRCRAFT_TAGS)).Select(element => element.Text).ToList();
        }
        private bool IsMatch(string expected, string actual)
        {
            return expected == actual;
        }

        private bool IsMatch(DateTime expected, DateTime actual)
        {
            return DateTime.Equals(expected, actual);
        }
        private bool AreListsMatching(List<string> expectedRoutes, List<string> actualRoutes)
        {
            if (expectedRoutes == null || actualRoutes == null)
            {
                return expectedRoutes == actualRoutes;
            }

            if (expectedRoutes.Count != actualRoutes.Count)
            {
                return false;
            }

            for (int i = 0; i < expectedRoutes.Count; i++)
            {
                if (expectedRoutes[i] != actualRoutes[i])
                {
                    return false;
                }
            }

            return true;
        }
        public List<string> GetLPCartsFiltred()
        {   WaitPageLoading();
            var listElement = _webDriver.FindElement(By.XPath(LP_CART_NAMES));
            var ordersText = listElement.Text.Trim();
            var ordersList = ordersText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(order => order.Trim()).ToList();
            return ordersList;
        }
        
        public string GetFlightCountNumber()
        {
            _flightsCount = WaitForElementIsVisible(By.XPath(FLIGHTS_COUNT));
            return _flightsCount.Text;
        }

        public bool VerifyLPCartAircraftInfo(string lpCartName, List<string> lpCartAircraft)
        {
            var generalInformationTab = WaitForElementIsVisible(By.XPath(GENERAL_INFO_TAB));
            generalInformationTab.Click();
            WaitPageLoading();
            var name = WaitForElementIsVisible(By.Id(LP_CART_NAME)).GetAttribute("value");

            return IsMatch(lpCartName, name) && AreListsMatching(lpCartAircraft, GetAircrafts());
        }

        public void CloseDatePicker()
        {
            _endDate = WaitForElementIsVisible(By.Id(END_DATE_FILTER));
            _endDate.SendKeys(Keys.Tab);
        }
    }
}

