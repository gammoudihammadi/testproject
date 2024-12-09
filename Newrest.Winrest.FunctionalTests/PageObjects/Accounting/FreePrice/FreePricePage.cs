using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Text;
using System.Threading;
using System.Web.UI.WebControls;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.FreePrice
{
    public class FreePricePage : PageBase
    {
        public FreePricePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________________ Constantes _________________________________________

        public const string FIRST_FREE_PRICE_NAME = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[2]";
        public const string FIRST_FREE_PRICE_FOREIGN_NAME = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[3]";
        public const string SEARCH_FILTER = "SearchPattern";
        public const string SHOW_ALL_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[3]";
        private const string NUMBER_FREE_PRICES = "/html/body/div[2]/div/div[2]/div/div[1]/h1/span";
        private const string FILTER_SITES = "SelectedSites_ms";
        private const string FILTER_SORT = "cbSortBy";
        private const string FILTER_CUSTOMER = "SelectedCustomers_ms";
        private const string COMBO_STATUS = "SelectedStatus_ms";
        public const string SHOW_ACTIVE = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[1]";
        public const string SHOW_ONLY_INACTIVE = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[2]";
        public const string CUSTOMERS_BTN = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[6]/div/div/div[1]/button";
        public const string SITES_BTN = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[5]/div/div/div[1]/button";
        private const string LIST_SITE_FILTER = "/html/body/div[10]/ul/li[*]/label/input";
        private const string LIST_CUSTOMER_FILTER = "/html/body/div[11]/ul/li[*]/label/input";
        public const string CUSTOMER_BY_TEXT = "/html/body/div[11]/ul/li[*]/label/span[text()='{0}']";
        public const string SITE_BY_TEXT = "/html/body/div[10]/ul/li[*]/label/span[text()='{0}']";

        public const string UNCHECK_ALL_CUSTOMERS = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        public const string UNCHECK_ALL_SITES = "/html/body/div[11]/div/ul/li[2]/a";
        public const string FREE_PRICE_NAMES = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr[*]/td[2]";
        private const string LIST_SITES_IN_GRID = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr[*]/td[5]";
        private const string LIST_CUSTOMERS_IN_GRID = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr[*]/td[4]";
        private const string RESET_FILTER = "reset-Filter";
        private const string LIST_SORT_FILTER = "cbSortBy";
        private const string SORT_BY_NAME = "//*[@id=\"cbSortBy\"]/option[2]";
        private const string SORT_BY_ID = "//*[@id=\"cbSortBy\"]/option[1]";
        private const string SORT_BY_ID_DESC = "//*[@id=\"cbSortBy\"]/option[3]";
        private const string EXTEND_MENU = "//*[@id=\"tabContentFreePricesContainer\"]/div[1]/div/div/button";
        private const string EXPORT_BUTTON = "//*[@id=\"btn-export-excel\"]";
        private const string IMPORT_BUTTON = "//*[@id=\"btn-Import\"]";
        private const string MASSIVE_DELETE_BUTTON = "//*[@id=\"tabContentFreePricesContainer\"]/div[1]/div/div/div/a[3]";

        private const string CHOOSE_FILE = "fileSent";
        private const string CHECK_FILE_BTN = "//*[@id=\"ImportFileForm\"]/div[3]/button[2]";

        private const string FREE_PRICES = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr[*]";



        // __________________________________________ Variables __________________________________________

        [FindsBy(How = How.XPath, Using = CHECK_FILE_BTN)]
        private IWebElement _checkFileBtn;

        [FindsBy(How = How.Id, Using = CHOOSE_FILE)]
        private IWebElement _chooseFile;

        [FindsBy(How = How.XPath, Using = FIRST_FREE_PRICE_NAME)]
        private IWebElement _firstFreePriceName;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        // __________________________________________ Méthodes ___________________________________________

        public enum FilterType
        {
            Search,
            Sites,
            ShowOnlyActive,
            ShowAll,
            ShowOnlyInactive,
            Customer,
            SortByName,
            SortById,
            SortByIdDescending,
        }

        public void Filter(FilterType filterType, object value, string sortBy = null)

        {
            switch (filterType)

            {
                case FilterType.Search:

                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.ShowAll:
                    var showAll = WaitForElementIsVisible(By.XPath(SHOW_ALL_FILTER));
                    showAll.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowOnlyActive:
                    var showonlyActive = WaitForElementIsVisible(By.XPath(SHOW_ACTIVE));
                    showonlyActive.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowOnlyInactive:
                    var showonlyInactive = WaitForElementIsVisible(By.XPath(SHOW_ONLY_INACTIVE));
                    showonlyInactive.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.Customer:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_CUSTOMER, (string)value));
                    break;

                case FilterType.Sites:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITES, (string)value));
                    break;

                case FilterType.SortByName:
                    var inputN = WaitForElementIsVisible(By.Id(LIST_SORT_FILTER));
                    inputN.Click();
                    var sortByName = WaitForElementIsVisible(By.XPath(SORT_BY_NAME));
                    sortByName.Click();
                    break;
                case FilterType.SortById:
                    var inputId = WaitForElementIsVisible(By.Id(LIST_SORT_FILTER));
                    inputId.Click();
                    var sortById = WaitForElementIsVisible(By.XPath(SORT_BY_ID));
                    sortById.Click();
                    break;
                case FilterType.SortByIdDescending:
                    var inputD = WaitForElementIsVisible(By.Id(LIST_SORT_FILTER));
                    inputD.Click();
                    var sortByIdDescending = WaitForElementIsVisible(By.XPath(SORT_BY_ID_DESC));
                    sortByIdDescending.Click();
                    break;
                default:
                    break;

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public object GetFilterValue(FilterType filterType)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    return _searchFilter.GetAttribute("value");

                case FilterType.ShowAll:
                    var _showAll = WaitForElementIsVisible(By.XPath(SHOW_ALL_FILTER));
                    return _showAll.Selected;

                case FilterType.ShowOnlyActive:
                    var _showOnlyActive = WaitForElementIsVisible(By.XPath(SHOW_ACTIVE));
                    return _showOnlyActive.Selected;

                case FilterType.ShowOnlyInactive:
                    var _showOnlyInactive = WaitForElementIsVisible(By.XPath(SHOW_ONLY_INACTIVE));
                    return _showOnlyInactive.Selected;
            }
            return null;
        }

        public FreePriceDetailsPage SelectFirstFreePrice()
        {
            _firstFreePriceName = WaitForElementIsVisible(By.XPath(FIRST_FREE_PRICE_NAME));
            _firstFreePriceName.Click();
            WaitForLoad();

            return new FreePriceDetailsPage(_webDriver, _testContext);
        }
        public string GetFirstFreePriceName()
        {
            _firstFreePriceName = WaitForElementIsVisible(By.XPath(FIRST_FREE_PRICE_NAME));
            var name =  _firstFreePriceName.Text;
            WaitForLoad();

            return name;
        }
        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();

            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //pas de date
            }
        }

        public string GetNumberInHeaderFreePrices()
        {
            var numberFreePrices = WaitForElementIsVisible(By.XPath(NUMBER_FREE_PRICES));
            return numberFreePrices.Text;
        }

        public bool VerifierShowAllFilter()
        {
            Filter(FilterType.ShowOnlyActive, true);
            var actifNumber = GetNumberInHeaderFreePrices();
            Filter(FilterType.ShowOnlyInactive, true);
            var inActifNumber = GetNumberInHeaderFreePrices();
            Filter(FilterType.ShowAll, true);
            var allNumber = GetNumberInHeaderFreePrices();

            if(int.Parse(allNumber) != (int.Parse(actifNumber) + int.Parse(inActifNumber)))
            {
                return false;
            }
            return true;
        }
        public IEnumerable<string> GetSites()
        {
            var listSitesInGrid = _webDriver.FindElements(By.XPath(LIST_SITES_IN_GRID));
            return listSitesInGrid.Select(p => p.Text);
        }
        public IEnumerable<string> GetCustomers()
        {
            var listCustomersInGrid = _webDriver.FindElements(By.XPath(LIST_CUSTOMERS_IN_GRID));
            return listCustomersInGrid.Select(p => p.Text);
        }

        public bool VerifyFilterSite(IEnumerable<string> listSitesInGrids, string site)
        {
            foreach (var siteInGrid in listSitesInGrids)
            {
                if (siteInGrid != site)
                {
                    return false;
                }
            }
            return true;
        }
        public bool VerifyFilterCustomer(IEnumerable<string> listCustomersInGrids, string site)
        {
            foreach (var customerInGrid in listCustomersInGrids)
            {
                if (customerInGrid != site)
                {
                    return false;
                }
            }
            return true;
        }
        public int GetNumberSelectedSiteFilter()
        {
            var listSitesSelectedFilters = _webDriver.FindElements(By.XPath(LIST_SITE_FILTER));
            var numberSitesSelectedSite = listSitesSelectedFilters
               .Where(p => p.Selected == true).Count();

            return numberSitesSelectedSite;
        }
        public int GetNumberSelectedCustomerFilter()
        {
            var listSelectedCustomerFilters = _webDriver.FindElements(By.XPath(LIST_CUSTOMER_FILTER));
            var numberSelectedCustomer = listSelectedCustomerFilters.Where(p => p.Selected).Count();

            return numberSelectedCustomer;
        }
        public bool IsSortedByName()
        {
            var ancientName = "";
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

            var elements = _webDriver.FindElements(By.XPath(FREE_PRICE_NAMES));

            foreach (var elm in elements)
            {

                if (elm.Text.CompareTo(ancientName) < 0)
                    return false;

                ancientName = elm.Text;

            }

            return true;

        }
        public bool VerifySortById()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var freePrices = _webDriver.FindElements(By.XPath(FREE_PRICES));
            foreach (var elm in freePrices)
            {

                string url = elm.GetAttribute("data-href");
                int queryIndex = url.IndexOf('?');
                if (queryIndex != -1)
                {
                    string queryString = url.Substring(queryIndex + 1);
                    var queryParams = System.Web.HttpUtility.ParseQueryString(queryString);
                    var idValue = queryParams["id"];
                    ids.Add(idValue);
                }
            }
            for (int i = 0; i < ids.Count - 1; i++)
            {
                if (int.Parse(ids[i]) > int.Parse(ids[i + 1]))
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public bool VerifySortByIdDescending()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var freePrices = _webDriver.FindElements(By.XPath(FREE_PRICES));
            foreach (var elm in freePrices)
            {

                string url = elm.GetAttribute("data-href");
                int queryIndex = url.IndexOf('?');
                if (queryIndex != -1)
                {
                    string queryString = url.Substring(queryIndex + 1);
                    var queryParams = System.Web.HttpUtility.ParseQueryString(queryString);
                    var idValue = queryParams["id"];
                    ids.Add(idValue);
                }
            }
            for (int i = 0; i < ids.Count - 1; i++)
            {
                if (int.Parse(ids[i]) < int.Parse(ids[i + 1]))
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public void Export()
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            var exportBtn = WaitForElementIsVisible(By.XPath(EXPORT_BUTTON));
            actions.MoveToElement(exportBtn).Perform();
            exportBtn.Click();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            ClickPrintButton();
            WaitForDownload();
            Close();
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
            Regex r = new Regex("[export?flights\\s\\d.-]", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);
            return m.Success;
        }

        public void MassiveDelete()
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            var massiveDeleteBtn = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_BUTTON));
            actions.MoveToElement(massiveDeleteBtn).Perform();
            massiveDeleteBtn.Click();

            WaitForLoad();
        }

        private const string SHOW_INACTIVE_SITES = "//*[@id=\"ShowInactiveSites\"]";
        private const string SHOW_INACTIVE_CUSTOMERS = "//*[@id=\"ShowInactiveCustomers\"]";
        private const string FREE_PRICE_NAME = "/html/body/div[3]/div/div/div[2]/div/form/div/div[1]/div/input";
        private const string FIRST_SITE_NAME = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr[1]/td[5]";
        private const string FIRST_CUSTOMER_NAME = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr[1]/td[4]";
        private const string COMBO_SITE = "SelectedSiteIds_ms";
        private const string COMBO_CUSTOMER = "SelectedCustomersForDeletion_ms";
        private const string SEARCH_BTN = "//*[@id=\"SearchFreePricesBtn\"]/span";
        private const string SELECT_ALL_BTN = "//*[@id=\"selectAll\"]/span";
        private const string DELETE_FREE_PRICE = "//*[@id=\"deleteFreePricesBtn\"]";
        private const string CONFIRM_DELETE_FREE_PRICE = "//*[@id=\"dataConfirmOK\"]";
        private const string OK_CONFIRM_DELETE_FREE_PRICE_MODAL = "closebtn";


        public void DeleteFreePrice(bool showInactiveSites, bool showInactiveCustomers, string freePriceame, string site, string customer, string status)
        {
            if(showInactiveSites)
            {
                var showInactiveSitesCheckbox = WaitForElementExists(By.XPath(SHOW_INACTIVE_SITES));
                showInactiveSitesCheckbox.Click();
            }
            if(showInactiveCustomers)
            {
                var showInactiveCustomersCheckbox = WaitForElementExists(By.XPath(SHOW_INACTIVE_CUSTOMERS));
                showInactiveCustomersCheckbox.Click();
            }

            var priceNameInput = WaitForElementExists(By.XPath(FREE_PRICE_NAME));
            priceNameInput.SetValue(ControlType.TextBox, freePriceame);
            WaitForLoad();

            ComboBoxSelectById(new ComboBoxOptions(COMBO_SITE, (string)site, false));

            ComboBoxSelectById(new ComboBoxOptions(COMBO_CUSTOMER, (string)customer, false));

            ComboBoxSelectById(new ComboBoxOptions(COMBO_STATUS, (string)status, false));

            var searchBtn = WaitForElementExists(By.XPath(SEARCH_BTN));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(searchBtn).Perform();
            searchBtn.Click();

            var selectAllBtn = WaitForElementExists(By.XPath(SELECT_ALL_BTN));
            actions.MoveToElement(selectAllBtn).Perform();
            selectAllBtn.Click();

            var deleteFreePriceBtn = WaitForElementExists(By.XPath(DELETE_FREE_PRICE));
            deleteFreePriceBtn.Click();

            var confirmDeleteFreePriceBtn = WaitForElementExists(By.XPath(CONFIRM_DELETE_FREE_PRICE));
            confirmDeleteFreePriceBtn.Click();

            if (IsDev())
            {
                var okCofirmationModal = WaitForElementIsVisible(By.Id(OK_CONFIRM_DELETE_FREE_PRICE_MODAL));
                okCofirmationModal.Click();

            }
            else
            {
                var okCofirmationModal = WaitForElementExists(By.XPath("//*/button[text()='Ok']"));
                okCofirmationModal.Click();
            }


        }

        public string GetFirstSite()
        {
            var siteName = WaitForElementExists(By.XPath(FIRST_SITE_NAME));
            return siteName.Text;
        }
        public string GetFirstCustomer()
        {
            var customerName = WaitForElementExists(By.XPath(FIRST_CUSTOMER_NAME));
            return customerName.Text;
        }

        public void ClickImport()
        {
            ClosePrintPopoverIfVisible();
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            WaitPageLoading();
            WaitForLoad();
            var importBtn = WaitForElementIsVisible(By.XPath(IMPORT_BUTTON));
            actions.MoveToElement(importBtn).Perform();
            importBtn.Click();
            //Léger Sleep sur l'import le temps que le fichier soit entièrement téléchargé
            //Thread.Sleep(2000);
            WaitPageLoading();
            WaitForLoad();

        }
        public void ClosePrintPopoverIfVisible()
        {
            var popoverPrintVisible = isElementVisible(By.XPath("//*[contains(@id,'popover')]"));
            if (popoverPrintVisible)
            {
                var btn = WaitForElementIsVisible(By.Id("header-print-button"));
                btn.Click();
            }
        }

        public bool ImportFile(string fullPath)
        {
            bool valueBool = true;
            try
            {
                CheckFile(fullPath);
                var btn_close = WaitForElementIsVisible(By.XPath("//*[@id=\"Cancel-btn\"]"));
                 btn_close.Click();
                WaitForLoad();
            }
            catch (Exception)
            {
                valueBool = false;
            }
            return valueBool;
        }
        public void CheckFile(string fullPath)
        {
            // Selection d'un fichier à importer
            _chooseFile = WaitForElementIsVisible(By.Id(CHOOSE_FILE));
            _chooseFile.SendKeys(fullPath);
            WaitForLoad();
            // Vérificaton du fichier
            _checkFileBtn = WaitForElementToBeClickable(By.XPath(CHECK_FILE_BTN));
            _checkFileBtn.Click();
            WaitForLoad();
        }
    }
}
