using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class SuppliersPage : PageBase
    {
        public SuppliersPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_______________________________________

        //Suppliuer
        private const string NEW_SUPPLIER = "New Supplier";
        private const string SUPPLIER_ITEMS_XPATH = "//*[@id=\"tableSuppliers\"]/tbody";

        
        //Export
        private const string EXPORT = "btn-export-suppliers-report";
        private const string NBOFITEMS = "btn-export-nb-items-report";
        private const string PRINT = "btn-print-suppliers-report";
        private const string IMPORT = "//*/a[text()='Import' and @role='button']";
        private const string DELIVERY_DAY = "btn-update-delivery-days";
        private const string DUPLICATE_SUPPLIER = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[10]";
        private const string DESACTIVATE_ITEMS = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[11]";


        //Items
        private const string FIRST_SUPPLIER_NAME = "//*[@id=\"tableSuppliers\"]/tbody/tr/td[3]";

        //Filter
        private const string RESET_FILTER_DEV = "ResetFilter";

        private const string FILTER_SEARCH_BY_NAME = "tbSearchPatternWithAutocomplete";
        private const string FIRST_RESULT_SEARCH = "//*[@id=\"formSearchSuppliers\"]/div[2]/span/div/div/div/strong[text()='{0}']";
        private const string FILTER_SORT_BY = "cbSortBy";

        private const string FILTER_SHOW_ALL_DEV = "ShowAll";
        private const string FILTER_SHOW_ONLY_ACTIVE_DEV = "ShowActiveOnly";
        private const string FILTER_SHOW_ONLY_INACTIVE_DEV = "ShowInactiveOnly";

        private const string FILTER_DELIVERY_DAYS = "SelectedDeliveryDays_ms";
        private const string SEARCH_DAY = "/html/body/div[10]/div/div/label/input";
        private const string UNSELECT_ALL_DAYS = "/html/body/div[10]/div/ul/li[2]/a";
        private const string FILTER_SITES = "SelectedSites_ms";
        private const string FILTER_SITES_ITEM_IN_SUPPLIER = "ItemsIndexModelSelectedSites_ms";

        private const string SEARCH_SITE = "/html/body/div[11]/div/div/label/input";
        private const string UNCHECK_ALL_SITES = "/html/body/div[11]/div/ul/li[2]/a";
        private const string INACTIVE2 = "//*[@id=\"tableSuppliers\"]/tbody/tr[*]/td[2]/img";
        private const string NAME = "//*[@id=\"tableSuppliers\"]/tbody/tr[*]/td[3]";

        private const string COUNTER = "//*[@id=\"div-body\"]/div/div[2]/div[1]/h1/span";
        private const string EXTEND_MENU = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/button";
        private const string DELIVERY_SITES = "btn-update-delivered-site";
        private const string COMBO_SITES = "SelectedDeliveredSites_ms";
        private const string COMBO_SUPPLIER = "SelectedSuppliers_ms";

        private const string COMBO_INPUT_SEARCH = "/html/body/div[13]/div/div/label/input";
        private const string FIRST_LINE = "/html/body/div[13]/ul/li[2]/label/input";
        private const string DELIVERED_SITES_SEARCH_BUTTON = "SearchSiteBtn";
        private const string UPDATE_BUTTON = "//*[@id=\"formSupplierDeliveredSites\"]/div[3]/button[1]";
        private const string FIRST_SUPPLIER = "//*/tr[contains(@data-href,'/Purchasing/Supplier/Details?id=')]";

        private const string SUPPLIER_NUMBER = "//*[@id=\"tableDeliveredSite\"]/tbody/tr[1]/td[3]";
        private const string FAST_COPY_PRICE = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[4]";
        private const string COPY_PRICE_SITES = "TargetedSites_ms";
        private const string COPY_PRICE_FIRST_SITE = "/html/body/div[13]/ul/li[1]/label/input";
        private const string COPY_PRICE_SUPPLIERS = "FilteredSuppliers_ms";
        private const string COPY_PRICE_FIRST_SUPPLIER = "/html/body/div[14]/ul/li[1]/label/input";
        private const string COPY_PRICE_UPDATE = "fastCopyPriceBtn";
        private const string SET_SITE = "/html/body/div[3]/div/div/div/div/form/div[2]/div[2]/div[2]/div/table/tbody/tr/td[3]/select";
        private const string FIRST_SITE = "/html/body/div[3]/div/div/div/div/form/div[2]/div[2]/div[2]/div/table/tbody/tr/td[3]/select/option[2]";
        private const string FIRST_SUPPLIER_CHECKBOX = "//*/td[contains(@class,'td-checkbox')]/input[1]";
        private const string MESSAGE_FAST_COPY_PRICE = "/html/body/div[3]/div/div/div/div[2]/div/h4";
        private const string SITES = "/html/body/div[11]/ul";

        private const string MONDAY = "//*/label[text()='Mon']";
        private const string TUESDAY = "//*/label[text()='Tue']";
        private const string WEDNESDAY = "//*/label[text()='Wed']";
        private const string THURSDAY = "//*/label[text()='Thu']";
        private const string FRIDAY = "//*/label[text()='Fri']";
        private const string SATURDAY = "//*/label[text()='Sat']";
        private const string SUNDAY = "//*/label[text()='Sun']";

        //Duplicate Supplier
        private const string FROM_SUPPLIER_DUPLICATE_SUPPLIER = "drop-down-supplier";
        private const string NEW_SUPPLIER_DUPLICATE_SUPPLIER = "NewSupplierName";
        private const string UPDATE_DUPLICATE_SUPPLIER = "duplicateSupplierBtn";

        // desactivate Item
        private const string COMBO_SITES_ON_DELIVERY_DAY = "SelectedDeliveredSites_ms";
        private const string COMBO_SITES_DESACTIVATE_ITEM = "SelectedSitesToUpdate_ms";

        private const string ITEM8GATEGORIE = "//*[@id=\"ItemFilterSelected\"]/option[1]";
        private const string CHECK_ALL_SITES = "/html/body/div[14]/div/ul/li[1]/a/span[2]";
        private const string COMBO_SUPPLIERS_DESACTIVATE_ITEM = "collapseSelectedSuppliersFilter";
        private const string CHECK_ALL_SUPPLIERS = "/html/body/div[15]/div/ul/li[1]/a/span[2]";
        private const string ITEMS  = "ItemFilterSelected";
        private const string ITEMS_DEACTIVATION_REPORT_TITLE = "//*[@id=\"formId\"]/div[1]/h4";
        private const string GOTOITEMSINSUPPLIERS = "/html/body/div[3]/div/ul/li[2]/a";
        private const string SETUNPERCHASABLE = "unpurchasable-items-btn";
        private const string NUMBERITEMS = "/html/body/div[4]/div/div/div[2]/div[1]/p";
        private const string CHECKALL = "/html/body/div[13]/div/ul/li[1]/a";
        private const string BTN_CLOSE_DELIVERY_SITE = "//*[@id=\"formSupplierDeliveredSites\"]/div[3]/button[2]";
        private const string PAGE_SIZE = "//*[@id=\"formSupplierDeliveredSites\"]//*[@id=\"page-size-selector\"]";
        private const string FILTER_SUPPLIERTYPE = "SelectedSupplierTypes_ms";

        //__________________________________Variables_______________________________________

        [FindsBy(How = How.XPath, Using = MONDAY)]
        private IWebElement _monday;
        [FindsBy(How = How.XPath, Using = CHECKALL)]
        private IWebElement _checkall;
        [FindsBy(How = How.XPath, Using = NUMBERITEMS)]
        private IWebElement _numberitems;
        [FindsBy(How = How.Id, Using = SETUNPERCHASABLE)]
        private IWebElement _setunperchsable;
        [FindsBy(How = How.XPath, Using = GOTOITEMSINSUPPLIERS)]
        private IWebElement _itemsinsupplier;
        [FindsBy(How = How.XPath, Using = TUESDAY)]
        private IWebElement _tuesday;

        [FindsBy(How = How.XPath, Using = WEDNESDAY)]
        private IWebElement _wednesday;

        [FindsBy(How = How.XPath, Using = THURSDAY)]
        private IWebElement _thursday;

        [FindsBy(How = How.XPath, Using = FRIDAY)]
        private IWebElement _friday;

        [FindsBy(How = How.XPath, Using = SATURDAY)]
        private IWebElement _saturday;

        [FindsBy(How = How.XPath, Using = SUNDAY)]
        private IWebElement _sunday;

        [FindsBy(How = How.XPath, Using = DUPLICATE_SUPPLIER)]
        private IWebElement _duplicate_supplier;

        [FindsBy(How = How.XPath, Using = DESACTIVATE_ITEMS)]
        private IWebElement _desactivateitem;


        [FindsBy(How = How.Id, Using = DELIVERY_DAY)]
        private IWebElement _deliveryday;

        [FindsBy(How = How.LinkText, Using = NEW_SUPPLIER)]
        private IWebElement _creatNewSupplierBtn;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = NBOFITEMS)]
        private IWebElement _nbOfItems;

        [FindsBy(How = How.Id, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.XPath, Using = IMPORT)]
        private IWebElement _import;

        [FindsBy(How = How.XPath, Using = FIRST_SUPPLIER_NAME)]
        private IWebElement _firstSupplierName;

        //__________________________________Filters_______________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.Id, Using = FILTER_SEARCH_BY_NAME)]
        private IWebElement _searchByName;

        [FindsBy(How = How.Id, Using = FILTER_SORT_BY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_ALL_DEV)]
        private IWebElement _showAllDev;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_ONLY_ACTIVE_DEV)]
        private IWebElement _showOnlyActiveDev;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_ONLY_INACTIVE_DEV)]
        private IWebElement _showOnlyInactiveDev;

        [FindsBy(How = How.Id, Using = FILTER_DELIVERY_DAYS)]
        private IWebElement _deliveryDays;

        [FindsBy(How = How.XPath, Using = SEARCH_DAY)]
        public IWebElement _searchDay;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_DAYS)]
        public IWebElement _unselectAllDays;

        [FindsBy(How = How.Id, Using = FILTER_SITES)]
        private IWebElement _sites;
        [FindsBy(How = How.Id, Using = FILTER_SITES_ITEM_IN_SUPPLIER)]
        private IWebElement _sites_in_supplieritem;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.XPath, Using = COUNTER)]
        private IWebElement _counter;

        [FindsBy(How = How.LinkText, Using = UNCHECK_ALL_SITES)]
        private IWebElement _uncheckAllSites;
        [FindsBy(How = How.LinkText, Using = FILTER_SUPPLIERTYPE)]
        private IWebElement _supplierType;

        //Duplicate Supplier

        [FindsBy(How = How.Id, Using = NEW_SUPPLIER_DUPLICATE_SUPPLIER)]
        private IWebElement _newsupplier;

        [FindsBy(How = How.Id, Using = UPDATE_DUPLICATE_SUPPLIER)]
        private IWebElement _updateduplicatesupplier;

        [FindsBy(How = How.XPath, Using = ITEM8GATEGORIE)]
        private IWebElement _itemcategorie;

        [FindsBy(How = How.XPath, Using = CHECK_ALL_SITES)]
        private IWebElement _checkAllSites;

        [FindsBy(How = How.XPath, Using = COMBO_SUPPLIERS_DESACTIVATE_ITEM)]
        private IWebElement _suppliers;

        [FindsBy(How = How.XPath, Using = CHECK_ALL_SUPPLIERS)]
        private IWebElement _checkAllSuppliers;

        [FindsBy(How = How.XPath, Using = ITEMS)]
        private IWebElement _items;
   
        [FindsBy(How = How.XPath, Using = ITEMS_DEACTIVATION_REPORT_TITLE)]
        private IWebElement _itemsDeactivationReportTitle;

        [FindsBy(How = How.XPath, Using = BTN_CLOSE_DELIVERY_SITE)]
        private IWebElement _btnCloseDeleviverySites;

        [FindsBy(How = How.XPath, Using = PAGE_SIZE)]
        private IWebElement _pageSize;
        public enum FilterType
        {
            Search,
            SortBy,
            ShowAll,
            ShowOnlyActive,
            ShowOnlyInactive,
            DeliveryDays,
            Site,
            DateFrom,
            DateTo,
            SupplierType
        }
        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchByName = WaitForElementIsVisibleNew(By.Id(FILTER_SEARCH_BY_NAME));
                    _searchByName.SetValue(ControlType.TextBox, value);

                    try
                    {
                        var firstResultSearch = _webDriver.FindElement(By.XPath(String.Format(FIRST_RESULT_SEARCH, value)));
                        firstResultSearch.Click();
                    }
                    catch
                    {
                        // item non trouvé
                    }
                    break;

                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisibleNew(By.Id(FILTER_SORT_BY));
                    _sortBy.SetValue(ControlType.DropDownList, value);
                    break;

                case FilterType.ShowAll:
                    _showAllDev = WaitForElementIsVisibleNew(By.Id(FILTER_SHOW_ALL_DEV));
                    _showAllDev.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.ShowOnlyActive:
                    _showOnlyActiveDev = WaitForElementIsVisibleNew(By.Id(FILTER_SHOW_ONLY_ACTIVE_DEV));
                    _showOnlyActiveDev.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.ShowOnlyInactive:
                    _showOnlyInactiveDev = WaitForElementIsVisibleNew(By.Id(FILTER_SHOW_ONLY_INACTIVE_DEV));
                    _showOnlyInactiveDev.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.DeliveryDays:
                    _deliveryDays = WaitForElementIsVisibleNew(By.Id(FILTER_DELIVERY_DAYS));
                    _deliveryDays.Click();

                    _unselectAllDays = WaitForElementIsVisibleNew(By.XPath(UNSELECT_ALL_DAYS));
                    _unselectAllDays.Click();

                    _searchDay = WaitForElementIsVisibleNew(By.XPath(SEARCH_DAY));
                    _searchDay.SetValue(ControlType.TextBox, value);

                    var valueToCheck = WaitForElementIsVisibleNew(By.XPath("//span[text()='" + value + "']"));
                    valueToCheck.Click();

                    _deliveryDays.Click();
                    break;

                case FilterType.Site:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITES, (string)value));
                    break;
                case FilterType.SupplierType:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SUPPLIERTYPE, (string)value));
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void FilterItemInSupplier(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchByName = WaitForElementIsVisible(By.Id(FILTER_SEARCH_BY_NAME));
                    _searchByName.SetValue(ControlType.TextBox, value);

                    try
                    {
                        var firstResultSearch = _webDriver.FindElement(By.XPath(String.Format(FIRST_RESULT_SEARCH, value)));
                        firstResultSearch.Click();
                    }
                    catch
                    {
                        // item non trouvé
                    }
                    break;

                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(FILTER_SORT_BY));
                    _sortBy.SetValue(ControlType.DropDownList, value);
                    break;

                case FilterType.ShowAll:
                    _showAllDev = WaitForElementIsVisible(By.Id(FILTER_SHOW_ALL_DEV));
                    _showAllDev.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.ShowOnlyActive:
                    _showOnlyActiveDev = WaitForElementIsVisible(By.Id(FILTER_SHOW_ONLY_ACTIVE_DEV));
                    _showOnlyActiveDev.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.ShowOnlyInactive:
                    _showOnlyInactiveDev = WaitForElementIsVisible(By.Id(FILTER_SHOW_ONLY_INACTIVE_DEV));
                    _showOnlyInactiveDev.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.DeliveryDays:
                    _deliveryDays = WaitForElementIsVisible(By.Id(FILTER_DELIVERY_DAYS));
                    _deliveryDays.Click();

                    _unselectAllDays = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_DAYS));
                    _unselectAllDays.Click();

                    _searchDay = WaitForElementIsVisible(By.XPath(SEARCH_DAY));
                    _searchDay.SetValue(ControlType.TextBox, value);

                    var valueToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    valueToCheck.Click();

                    _deliveryDays.Click();
                    break;

                case FilterType.Site:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITES_ITEM_IN_SUPPLIER, (string)value));
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilter()
        {
            _resetFilterDev = WaitForElementIsVisible(By.Id(RESET_FILTER_DEV));
            _resetFilterDev.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //pas de FilterType.DateTo implémenté dans le switch case
                //pas de DateTo dans l'IHM
            }
        }
        public void GoToItemsInSupplier()
        {
            _itemsinsupplier = WaitForElementIsVisible(By.XPath(GOTOITEMSINSUPPLIERS));
            _itemsinsupplier.Click();
            WaitForLoad();
    
        }
        public void ClickSetUnperchsableItem()
        {
            _setunperchsable = WaitForElementIsVisible(By.Id(SETUNPERCHASABLE));
            _setunperchsable.Click();
            WaitForLoad();

        }
        public bool CheckStatus()
        {

            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            var elements = _webDriver.FindElements(By.XPath(INACTIVE2));

            if (elements.Count() == 0)
            {
                return true;
            }
            else
            {
                return false;
            }

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

            var elements = _webDriver.FindElements(By.XPath(NAME));

            foreach (var elm in elements)
            {

                if (elm.Text.CompareTo(ancientName) < 0)
                    return false;

                ancientName = elm.Text;

            }

            return true;

        }

        public bool IsListSorted(List<string> list)
        {
            bool valueBool = true;

            if (list != null && list.Count > 0)
            {
                string initialValue = list[0];

                foreach (string number in list)
                {
                    if (String.Compare(initialValue, number) > 1)
                    {
                        valueBool = false;
                        break;
                    }

                    initialValue = number;
                }
            }
            else
            {
                valueBool = false;
            }

            return valueBool;
        }

        // _____________________________________ Utilitaires __________________________________________

        public SupplierItem SelectFirstItem()
        {
            _firstSupplierName = WaitForElementIsVisibleNew(By.XPath(FIRST_SUPPLIER_NAME));
            _firstSupplierName.Click();
            WaitForLoad();

            return new SupplierItem(_webDriver, _testContext);
        }
        public IList<IWebElement> WaitForElementsAreVisible(By by, int timeoutInSeconds = 10)
        {
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.VisibilityOfAllElementsLocatedBy(by));
        }

        public SupplierItem SelectRandomItem()
        {
            // Get all supplier elements
            var supplierElements = WaitForElementsAreVisible(By.XPath(SUPPLIER_ITEMS_XPATH));

            // Check if the list is not empty
            if (supplierElements == null || supplierElements.Count == 0)
            {
                throw new NoSuchElementException("No supplier items found.");
            }

            // Select a random supplier item
            var random = new Random();
            var randomSupplierElement = supplierElements[random.Next(supplierElements.Count)];
            randomSupplierElement.Click();
            WaitForLoad();

            // Return the selected SupplierItem
            return new SupplierItem(_webDriver, _testContext);
        }

        public SupplierGeneralInfoTab SelectFirstSupplier()
        {
            _firstSupplierName = WaitForElementIsVisible(By.XPath(FIRST_SUPPLIER_NAME));
            _firstSupplierName.Click();
            WaitForLoad();

            return new SupplierGeneralInfoTab(_webDriver, _testContext);
        }

        public string GetFirstSupplierName()
        {
            if (isElementVisible(By.XPath(FIRST_SUPPLIER_NAME)))
            {
                _firstSupplierName = WaitForElementIsVisible(By.XPath(FIRST_SUPPLIER_NAME));
                return _firstSupplierName.Text;
            }
            else
            {
                return "";
            }

        }

        public SupplierCreateModalPage SupplierCreatePage()
        {
            ShowPlusMenu();

            _creatNewSupplierBtn = WaitForElementIsVisibleNew(By.LinkText(NEW_SUPPLIER));
            _creatNewSupplierBtn.Click();

            return new SupplierCreateModalPage(_webDriver, _testContext);
        }
        public void Export(bool newVersionPrint)
        {
            ShowExtendedMenu();
            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();

            if (newVersionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            }

            WaitForDownload();
            Close();
        }

        public void NberOfitems()
        {
            ShowExtendedMenu();
            _nbOfItems = WaitForElementIsVisible(By.Id(NBOFITEMS));
            _nbOfItems.Click();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            WaitForDownload();

            Close();
        }

        public void DeliveryDay()
        {
            ShowExtendedMenu();
            _deliveryday = WaitForElementIsVisible(By.Id(DELIVERY_DAY));
            _deliveryday.Click();
            WaitForLoad();
        }

        public PrintReportPage Print(bool newVersionPrint)
        {
            ShowExtendedMenu();
            _print = WaitForElementIsVisible(By.Id(PRINT));
            _print.Click();

            if (newVersionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            Close();

            return new PrintReportPage(_webDriver, _testContext);
        }


        public bool Import(FileInfo file)
        {
            ShowExtendedMenu();
            _import = WaitForElementIsVisible(By.XPath(IMPORT));
            _import.Click();

            var inputFile = WaitForElementIsVisible(By.Id("fileSent"));
            inputFile.SendKeys(file.FullName);
            var checkFile = WaitForElementIsVisible(By.XPath("//*[@id='form-import']/div[3]/button[2]"));
            checkFile.Click();
            WaitForLoad();

            var importFile = WaitForElementIsVisible(By.XPath("//*/button[text()='Import File']"));
            importFile.Click();
            WaitPageLoading();

            bool error = isElementVisible(By.XPath("//*/div[@class='errors-panel']/p/b[contains(text(),'Some error(s) have occurred')]"));
            var close = WaitForElementIsVisible(By.XPath("(//*/button[text()='Close'])[1]"));
            close.Click();
            WaitForLoad();

            return !error;
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles, bool parExtension = false)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                if (parExtension)
                {
                    var x = file.Name;
                    if (file.Name.EndsWith(".xlsx"))
                    {
                        correctDownloadFiles.Add(file);
                    }
                }
                else
                {
                    //  Test REGEX
                    if (IsSupplierExcelFileCorrect(file.Name))
                    {
                        correctDownloadFiles.Add(file);
                    }
                }
            }
            if (correctDownloadFiles.Count <= 0)
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

        public bool IsSupplierExcelFileCorrect(string filePath)
        {
            // "export-suppliers -.xlsx";
            string re1 = "((?:[a-z][a-z]+))";   // Word 1
            string re2 = "(.)"; // Any Single Character 1
            string re3 = "((?:[a-z][a-z]+))";   // Word 2
            string re4 = "(\\s+)";  // White Space 1
            string re5 = "(.)"; // Any Single Character 2
            string re6 = "(.)"; // Any Single Character 3
            string re7 = "((?:[a-z][a-z]+))";   // Word 3

            Regex r = new Regex(re1 + re2 + re3 + re4 + re5 + re6 + re7, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public bool IsSupplier(List<string> suppliers)
        {
            bool isSupplier = false;

            foreach (var supplier in suppliers)
            {
                Filter(SuppliersPage.FilterType.Search, supplier);
                WaitForLoad();

                _counter = WaitForElementIsVisible(By.XPath(COUNTER));

                int numberOfSupplier = Convert.ToInt32(_counter.Text);

                if (numberOfSupplier > 0)
                {
                    isSupplier = true;
                }
            }

            return isSupplier;
        }
        public void ClickDeliverySites()
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU));
            extendMenu.Click();
            var deliverySites = WaitForElementIsVisible(By.Id(DELIVERY_SITES));
            deliverySites.Click();
        }
        public void SetDeliveredSiteOnSupplier(string site, string supplierName)
        {
            // Select the site from the ComboBox and initiate the search.
            ComboBoxSelectById(new ComboBoxOptions(COMBO_SITES, site, false));
            var searchBtt = WaitForElementIsVisible(By.Id(DELIVERED_SITES_SEARCH_BUTTON));
            searchBtt.Click();
            WaitForLoad();
            PageSizeModalSuppliersPage("100");

            int pageNumber = 1;
            bool checkboxSet = false;

            while (!checkboxSet)
            {
                // Find the supplier checkbox on the current page.
                var supplierCheckbox = _webDriver.FindElements(By.XPath($"//*[@id=\"tableDeliveredSite\"]/tbody/tr[td[2][contains(text(),'{supplierName.ToLower()}')]]/td[1]"));

                if (supplierCheckbox.Count > 0 && supplierCheckbox[0].Displayed)
                {
                    // If the checkbox is found and is not selected, select it.
                    if (!supplierCheckbox[0].Selected)
                    {
                        supplierCheckbox[0].Click();
                        checkboxSet = true;

                        // Click the update button to save changes.
                        var blueUpdate = WaitForElementIsVisible(By.XPath(UPDATE_BUTTON));
                        blueUpdate.Click();
                        WaitForLoad();
                        CloseDeliverySites();
                    }
                    else
                    {
                        checkboxSet = true;
                    }
                }
                else
                {
                    // If checkbox not found on current page, go to the next page if it exists.
                    pageNumber++;
                    GoToNextPage(pageNumber.ToString());
                    WaitPageLoading();
                    WaitPageLoading();
                }
            }
        }

        public void SetSiteOnDeliveryDay(string site)
        {
            ComboBoxSelectById(new ComboBoxOptions(COMBO_SITES_ON_DELIVERY_DAY, site, false));

            WaitForLoad();
        }
        public void SetSiteDesactivateItem(string site)
        {
            ComboBoxSelectById(new ComboBoxOptions(COMBO_SITES_DESACTIVATE_ITEM, site, false));
            WaitForLoad();
        }

        public void SetSupplierOnDeliveryDay(string site)
        {
            ComboBoxSelectById(new ComboBoxOptions(COMBO_SUPPLIER, site, false));
            WaitForLoad();
        }

        public void CompletDeliveriesDaysAndComment()
        {
            // Select days for ALL SITES
            _monday = WaitForElementIsVisible(By.XPath(MONDAY));
            _monday.Click();

            _tuesday = WaitForElementIsVisible(By.XPath(TUESDAY));
            _tuesday.Click();

            _wednesday = WaitForElementIsVisible(By.XPath(WEDNESDAY));
            _wednesday.Click();

            _thursday = WaitForElementIsVisible(By.XPath(THURSDAY));
            _thursday.Click();

            _friday = WaitForElementIsVisible(By.XPath(FRIDAY));
            _friday.Click();

            _saturday = WaitForElementIsVisible(By.XPath(SATURDAY));
            _saturday.Click();
            WaitForLoad();

            //Ajout Thread.sleep() avant de passer à la prochaine étape
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public string[] SelectedDeliveriesDays()
        {
            string[] deliveryday = new string[7];

            // Select days for ALL SITES
            _monday = WaitForElementIsVisible(By.XPath(MONDAY));
            _monday.Click();
            deliveryday[0] = _monday.GetCssValue("border"); //1px solid rgb(238, 238, 238)

            _tuesday = WaitForElementIsVisible(By.XPath(TUESDAY));
            _tuesday.Click();
            deliveryday[1] = _tuesday.GetCssValue("border");

            _wednesday = WaitForElementIsVisible(By.XPath(WEDNESDAY));
            _wednesday.Click();
            deliveryday[2] = _wednesday.GetCssValue("border");

            _thursday = WaitForElementIsVisible(By.XPath(THURSDAY));
            _thursday.Click();
            deliveryday[3] = _thursday.GetCssValue("border");

            _friday = WaitForElementIsVisible(By.XPath(FRIDAY));
            //_friday.Click();
            deliveryday[4] = _friday.GetCssValue("border");

            _saturday = WaitForElementIsVisible(By.XPath(SATURDAY));
            //_saturday.Click();
            deliveryday[5] = _saturday.GetCssValue("border");

            _sunday = WaitForElementIsVisible(By.XPath(SUNDAY));
            //_sunday.Click();
            deliveryday[6] = _sunday.GetCssValue("border");

            WaitForLoad();
            //Ajout Thread.sleep() avant de passer à la prochaine étape
            Thread.Sleep(2000);
            WaitForLoad();

            CloseDeliveryDays();

            return deliveryday;

        }

        public void CloseDeliveryDays()
        {
            var close = WaitForElementIsVisible(By.XPath("//*[@id=\"formSupplierDeliveredSites\"]/div[3]/button"));
            close.Click();
            WaitForLoad();
        }

        public void ClickCopyPrice()
        {
            ShowExtendedMenu();
            var fastCopyPrice = WaitForElementIsVisible(By.XPath(FAST_COPY_PRICE));
            fastCopyPrice.Click();
        }
        public void CopyPriceSelectFirstSite()
        {
            ComboBoxSelectById(new ComboBoxOptions(COPY_PRICE_SITES, "ACE - ACE", false));
        }
        public void CopyPriceSelectFirstSupplier(string suppliers)
        {
            
            ComboBoxSelectById(new ComboBoxOptions(COPY_PRICE_SUPPLIERS, suppliers, false));

            var setSite = WaitForElementExists(By.XPath(SET_SITE));
            setSite.SetValue(ControlType.DropDownList, "MAD");
            WaitForLoad();
        }
        public void CopyPriceUpdate()
        {
            string suppliers = "ACEITES CONDE DE TORREPALMA, SCA";
            ClickCopyPrice();
            CopyPriceSelectFirstSite();
            CopyPriceSelectFirstSupplier(suppliers);
            var update = WaitForElementIsVisible(By.Id(COPY_PRICE_UPDATE));
            update.Click();
        }
        public bool VerifyFastCopyPrice()
        {
            var message = WaitForElementIsVisible(By.XPath(MESSAGE_FAST_COPY_PRICE));
            if (message.Text == "Import done!")
                return true;
            return false;
        }
        public bool verifyDonneeExcel(IEnumerable<string> list1, IEnumerable<string> list2)
        {
            foreach (var item in list2)
            {
                if (!list1.FirstOrDefault().Contains(item))
                {
                    return false;
                }
            }
            return true;
        }
        public string GetNameSearched()
        {
            _searchByName = WaitForElementExists(By.Id(FILTER_SEARCH_BY_NAME));
            return _searchByName.Text;
        }
        public string GetSortTypeSelected()
        {
            _sortBy = WaitForElementExists(By.Id(FILTER_SORT_BY));
            return _sortBy.GetAttribute("value");
        }
        public string GetNumItemsonPopum()
        {
            _numberitems = WaitForElementExists(By.XPath(NUMBERITEMS));
            string message = _numberitems.Text;
            var match = Regex.Match(message, @"All of (\d+) item\(s\) have been successfully set unpurchasable\s*!");
            if (match.Success)
            {
                return match.Groups[1].Value; 
            }

            return "0"; 
        }
        public string GetActiveOrInactiveChecked()
        {
            _showAllDev = WaitForElementIsVisible(By.Id(FILTER_SHOW_ALL_DEV));
            _showOnlyActiveDev = WaitForElementIsVisible(By.Id(FILTER_SHOW_ONLY_ACTIVE_DEV));
            _showOnlyInactiveDev = WaitForElementIsVisible(By.Id(FILTER_SHOW_ONLY_INACTIVE_DEV));
            if (_showAllDev.Selected)
            {
                return _showAllDev.GetAttribute("value");
            }
            else if (_showOnlyActiveDev.Selected)
            {
                return _showOnlyActiveDev.GetAttribute("value");
            }
            else if (_showOnlyInactiveDev.Selected)
            {
                return _showOnlyInactiveDev.GetAttribute("value");
            }

            return null;
        }
        public string GetDeliveryDaySelected()
        {
            _deliveryDays = WaitForElementExists(By.Id(FILTER_DELIVERY_DAYS));
            return _deliveryDays.Text;
        }
        public string GetSiteSelected()
        {
            _sites = WaitForElementExists(By.Id(FILTER_SITES));
            return _sites.Text;
        }
        public string GetSupplierTypesSelected()
        {
            _deliveryDays = WaitForElementExists(By.Id(FILTER_DELIVERY_DAYS));
            return _deliveryDays.Text;
        }

        public bool IsPageItemsCountEqualsTo50()
        {
            var list = _webDriver.FindElements(By.XPath("//*[@id=\"tableSuppliers\"]/tbody/tr[*]"));

            return list.Count == 50;
        }

        public bool IsPageItemsCountEqualsTo100()
        {
            var list = _webDriver.FindElements(By.XPath("//*[@id=\"tableSuppliers\"]/tbody/tr[*]"));

            return list.Count == 100;
        }
        public bool IsSuppliersEqualsToSizePage(int sizePage)
        {
            var suppliers = _webDriver.FindElements(By.XPath("/html/body/div[2]/div/div[2]/div[2]/div/table/tbody/tr[*]"));
            return suppliers.Count == sizePage;
        }

        public void ClickUNPURCHASABLEITEM()
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU));
            extendMenu.Click();
            ShowExtendedMenu();
            var deliverySites = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[7]"));
            deliverySites.Click();
        }

        public void SetUnpurchasableItem()
        {
            ComboBoxSelectById(new ComboBoxOptions("SelectedSitesToUpdate_ms", "MAD", false));

            ComboBoxSelectById(new ComboBoxOptions("SelectedSuppliers_ms", "SupplierTestDeleteItem", false));

            var setButton = WaitForElementIsVisible(By.Id("last"));
            setButton.Click();
            var okButton = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div[2]/button"));
            okButton.Click();
        }

        public SupplierItem ClickAndGoFirstSupplier()
        {
            _firstSupplierName = WaitForElementIsVisible(By.XPath(FIRST_SUPPLIER_NAME));
            _firstSupplierName.Click();
            WaitForLoad();
            return new SupplierItem(_webDriver, _testContext);
        }
        public void DuplicateSupplier()
        {
            ShowExtendedMenu();
            _duplicate_supplier = WaitForElementIsVisible(By.XPath(DUPLICATE_SUPPLIER));
            _duplicate_supplier.Click();
            WaitForLoad();
        }
        public void SetFromSupplierDuplicateSupplier(string fromsupplier)
        {
            var _site = WaitForElementIsVisible(By.Id(FROM_SUPPLIER_DUPLICATE_SUPPLIER));
            _site.SetValue(ControlType.DropDownList, fromsupplier);
            WaitForLoad();
        }
        public void SetNewSupplierDuplicateSupplier(string newsupplier)
        {
            _newsupplier = WaitForElementIsVisible(By.Id(NEW_SUPPLIER_DUPLICATE_SUPPLIER));
            _newsupplier.SetValue(ControlType.TextBox, newsupplier);
            WaitForLoad();
        }

        public void UncheckEDI()
        {
            var removeEDI = WaitForElementExists(By.Id("CopyFromEDI"));
            removeEDI.SetValue(ControlType.CheckBox, false);
            WaitForLoad();
        }

        public SupplierItem ClickButtonUpdate()
        {
            _updateduplicatesupplier = WaitForElementIsVisible(By.Id(UPDATE_DUPLICATE_SUPPLIER));
            _updateduplicatesupplier.Click();
            WaitPageLoading();
            WaitForLoad();
            return new SupplierItem(_webDriver, _testContext);
        }

        public void DesactivateItemr()
        {
            ShowExtendedMenu();
            _desactivateitem = WaitForElementIsVisible(By.XPath(DESACTIVATE_ITEMS));
            _desactivateitem.Click();
            WaitForLoad();
        }

        public void ClicktCategorieItem()
        {
            _itemcategorie = WaitForElementExists(By.XPath(ITEM8GATEGORIE));
            _itemcategorie.Click();
            WaitForLoad();
        }
        public void ClickDeleteDesactivateItem()
        {
            WaitForLoad();
            var delete = WaitForElementIsVisible(By.Id("btnSubmitId"));
            delete.Click();
            WaitPageLoading();
            WaitPageLoading();
            WaitPageLoading();
            Thread.Sleep(20000);


        }

        public void ClickSaveDesactivateItem()
        {
            WaitForLoad();
            var delete = WaitForElementIsVisible(By.XPath("//*[@id=\"formId\"]/div[3]/button[2]"));
            delete.Click();
            WaitForLoad();
        }

        public int GetSuppliersCount()
        {
            var suppliers = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div[1]/h1/span"));
            var count = int.Parse(suppliers.Text);
            return count;
        }
        public int GetItemsInSupplierCount()
        {

            var list = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div/div/div/div[2]/div/div[2]/div/div[*]/div[1]/div/div[2]/table/tbody/tr"));

            return list.Count();
        }
        public void Back_To_Navigate(int i)
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[i]);

            WaitForLoad();
        }

        public void Checkall()
        {
            _checkall = WaitForElementIsVisible(By.XPath(CHECKALL));
            _checkall.Click();
            WaitForLoad();

        }



        public void SetSiteDesactivateItem_Chceck_All()
        {

            var sites = WaitForElementIsVisible(By.Id(COMBO_SITES_DESACTIVATE_ITEM));
            sites.Click();
            var check_all_btn = WaitForElementIsVisible(By.XPath(CHECK_ALL_SITES));
            check_all_btn.Click();
            sites.Click();
            WaitForLoad();
        }
        public void SetSUPPLIERSDesactivateItem_Chceck_All()
        {
            var sites = WaitForElementIsVisible(By.Id(COMBO_SUPPLIERS_DESACTIVATE_ITEM));
            sites.Click();
            var check_all_btn = WaitForElementIsVisible(By.XPath(CHECK_ALL_SUPPLIERS));
            check_all_btn.Click();
            sites.Click();
            WaitForLoad();
        }
        public void SetIteams_DeactivateItems ( string ItemFiltered)
        {
            switch (ItemFiltered)
            {
                case "IsActifAndNotPurchasable":
                    DropdownListSelectById(ITEMS, ItemFiltered);
                  
                    WaitForLoad();
                    break;

                default:
                 
                    break;
            }
        }
        public bool ItemsDeactivationReport_IsVisible()
        {
            WaitPageLoading();

            var title = WaitForElementIsVisible(By.XPath(ITEMS_DEACTIVATION_REPORT_TITLE)); 
            if ( title.Text == "Items deactivation report")
                return true;
            else 
                return false;

        }
        //On peut utiliser la methode PageSize de PageBase
        //public new void PageSize(string size)
        //{

        //    var PageSizeDdl = _webDriver.FindElement(By.XPath("/html/body/div[3]/div/div/div/div/form/div[2]/div[2]/div/div/nav/select"));
        //    //PageSizeDdl.Click();
        //    PageSizeDdl.SetValue(ControlType.DropDownList, size);
        //    WaitForLoad();
        //}
        public bool VerifySite(string site)
        {
            var sites = _webDriver.FindElements(By.XPath(SITES));
            foreach (var s in sites)
            {
                if (s.Text.Trim().Contains(site))
                    return false;
            }
            return true;
        }
        public void GoToNextPage(string pageNumber)
        {
            var pageNumbers = _webDriver.FindElements(By.XPath("//*[@id=\"list-sites-delivered\"]/nav/ul/li[*]/a"));
            foreach (var page in pageNumbers)
            {
                if (page.Text == pageNumber)
                {
                    string xpath = "//*[@id=\"list-sites-delivered\"]/nav/ul/li[*]/a[text()=" + pageNumber + "]";
                    var pageBtn = WaitForElementIsVisible(By.XPath(xpath));
                    pageBtn.Click();
                }
            }

        }
        public void CloseDeliverySites()
        {
            var close = WaitForElementIsVisible(By.XPath(BTN_CLOSE_DELIVERY_SITE));
            close.Click();
            WaitForLoad();
        }
        public void PageSizeModalSuppliersPage(string size)
        {
            if (size == "1000")
            {   // Test
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                js.ExecuteScript("$('#" + PAGE_SIZE + "').append($('<option>', {value: 1000,text: '1000'}),'');");
            }

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            try
            {
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(PAGE_SIZE)));
            }
            catch
            {
                // tableau vide : pas de PageSize
                return;
            }
            _pageSize = WaitForElementExists(By.XPath(PAGE_SIZE));
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_pageSize).Perform();
            _pageSize = WaitForElementExists(By.XPath(PAGE_SIZE));
            _pageSize.SetValue(ControlType.DropDownList, size);

            WaitPageLoading();
            WaitForLoad();
            // pour écran plus petit que 8 lignes affiché
            PageUp();
        }
    }
}
