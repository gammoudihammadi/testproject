using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using Limilabs.Client.IMAP;
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
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item
{
    public class MassiveDeleteModal : PageBase
    {
        public MassiveDeleteModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        private const string SUPPLIER = "SelectedSupplierIds_ms";
        private const string SEARCH_BTN = "SearchItemsBtn";
        private const string UNSELECT_ALL_SUPPLIERS = "/html/body/div[17]/div/ul/li[2]/a/span[2]";
        private const string SELECT_ALL_SUPPLIERS = "/html/body/div[17]/div/ul/li[1]/a/span[2]";
        private const string SEARCH_BY_NAME = "SearchPattern";
        private const string SHOW_INACTIVE_SUPPLIERS = "//*[@id=\"formMassiveDeleteItem\"]/div/div[2]/div[2]/div[2]/div";
        private const string MASSIVE_DELETE_SELECTALL_BUTTON = "selectAll";
        private const string MASSIVE_DELETE_UNSELECTALL_BUTTON = "unselectAll";
        private const string MASSIVE_DELETE_ITEM_SELECT_COUNT = "itemCount";
        private const string DDL_PAGESIZE = "//*[@id=\"div-itemResultTable\"]/descendant::select[contains(@id, 'page-size-selector')]";
        private const string MASSIVE_DELETE_PAGESIZE = "/html/body/div[3]/div/div/div[2]/div/form/div/div[6]/div/div/nav/select";
        private const string SITE_FILTER = "SelectedSiteIds_ms";
        private const string UNSELECT_ALL_SITES = "/html/body/div[16]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_SITE = "/html/body/div[16]/div/div/label/input";
        private const string PAGESIZE = "/html/body/div[3]/div/div/div[2]/div/form/div/div[6]/div/div/nav/select";
        private const string ALL_SITES = "//*[@id=\"tableItems\"]/tbody/tr[*]/td[4]";
        private const string CLICK_SITE = "//*[@id=\"tableItems\"]/thead/tr/th[4]/span/a";
        private const string SITES_SELECT_ALL = "/html/body/div[16]/div/ul/li[1]/a/span[2]";
        private const string MASSIVE_ITEMS = "//*[@id=\"tableItems\"]/tbody/tr[*]/td[2]";
        private const string MASSIVE_DELETE_SORT_BY_ITEMS = "//*[@id=\"tableItems\"]/thead/tr/th[2]/span/a";
        private const string CLICK_STATUS = "//*[@id=\"tableItems\"]/thead/tr/th[3]/span/a";
        private const string MASSIVE_DELETE_SELECT_CHECKBOX = "item_IsSelected";
        private const string MASSIVE_DELETE_SERVICE_DELETE_BTN = "deleteItemBtn";
        private const string MASSIVE_DELETE_SERVICE_CONFIRM_BTN = "dataConfirmOK";
        private const string MASSIVE_DELETE_SERVICE_DELETEOK_BTN = "/html/body/div[3]/div/div/div[3]/button";
        private const string DELETE_BUTTON = "deleteItemBtn";
        private const string CONFIRM_DELETE = "dataConfirmOK";
        private const string CLOSE_POPUP = "//*[@id=\"modal-1\"]/div[3]/button";
        private const string CLICK_SUPPLIER = "//*[@id=\"tableItems\"]/thead/tr/th[5]/span/a";
        private const string COUNT_SELECTED_MENUS = "menuCount";


        private const string SUPPLIER_FILTER = "/html/body/div[3]/div/div/div[2]/div/form/div/div[2]/div[2]/div[1]/div/div/div/div[1]/button";
        private const string SUPLLIER_NAME = "//*[@id=\"tableItems\"]/tbody/tr[1]/td[5]";
     
        private const string SHOW_INACTIVE_SITE_CHECKBOX = "//*[@id=\"formMassiveDeleteItem\"]/div/div[2]/div[1]/div[2]/div";
        private const string SITES_COMBOBOX = "SelectedSiteIds_ms";
        private const string FIRST_ITEM_SITE_COMBOBOX = "/html/body/div[18]/ul/li[2]/label/span";
        private const string SECOND_ITEM_SITE_COMBOBOX = "/html/body/div[16]/ul/li[6]/label/span";
        private const string SITE_INPUT_COMBOBOX = "/html/body/div[18]/div/div/label/input";
        private const string SUPPLIER_SELECT_ALL = "/html/body/div[17]/div/ul/li[1]/a/span[2]";
        private const string SEARCH_SUPPLIER = "/html/body/div[17]/div/div/label/input";
        private const string SUPPLIER_UNSELECT_ALL = "/html/body/div[17]/div/ul/li[2]/a/span[2]";
        private const string SITE_INPUT_COMBO = "/html/body/div[17]/div/div/label/input";
        private const string FIRST_ITEM_SITE_COMBO = "/html/body/div[17]/ul/li[2]/label/span";
        private const string SHOW_INACTIVE_SITES = "//*[@id=\"formMassiveDeleteItem\"]/div/div[2]/div[1]/div[2]/div";
        private const string FILTRED_SITES_LIST = "/html/body/div[17]/ul";
        private const string INPUT_SEARCH = "/html/body/div[17]/div/div/label/input"; 

        [FindsBy(How = How.Id, Using = DELETE_BUTTON)]
        public IWebElement _deleteButton;
        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        public IWebElement _confirmDelete;
        [FindsBy(How = How.XPath, Using = CLOSE_POPUP)]
        public IWebElement _closePopup;
        [FindsBy(How = How.XPath, Using = SHOW_INACTIVE_SITE_CHECKBOX)]
        public IWebElement _showInactiveSiteCheckbox;
        [FindsBy(How = How.XPath, Using = SITE_INPUT_COMBOBOX)]
        public IWebElement _siteInputCombobox;
        [FindsBy(How = How.Id, Using = SITES_COMBOBOX)]
        public IWebElement _sitesCombobox;
        [FindsBy(How = How.XPath, Using = FIRST_ITEM_SITE_COMBOBOX)]
        public IWebElement _firstSiteCheckbox;
        [FindsBy(How = How.XPath, Using = SECOND_ITEM_SITE_COMBOBOX)]
        public IWebElement _secondSiteCheckbox;



        [FindsBy(How = How.Id, Using = SUPPLIER)]
        private IWebElement _supplier;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SUPPLIERS)]
        private IWebElement _unselectAllSuppliers;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_SUPPLIERS)]
        private IWebElement _selectAllSuppliers;

        [FindsBy(How = How.Id, Using = SEARCH_BY_NAME)]
        private IWebElement _searchByNameFilter;

        [FindsBy(How = How.Id, Using = SHOW_INACTIVE_SUPPLIERS)]
        private IWebElement _showinactivesuppliers;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SITES)]
        private IWebElement _unselectAllSites;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.XPath, Using = PAGESIZE)]
        private IWebElement _pageSize;

        [FindsBy(How = How.XPath, Using = CLICK_SITE)]
        private IWebElement _selectSite;

        [FindsBy(How = How.XPath, Using = SITES_SELECT_ALL)]
        private IWebElement _siteSelectAll;

        [FindsBy(How = How.XPath, Using = CLICK_STATUS)]
        private IWebElement _selectStatus;
        [FindsBy(How = How.XPath, Using = CLICK_SUPPLIER)]
        private IWebElement _selectSupplier;
        [FindsBy(How = How.XPath, Using = SUPPLIER_SELECT_ALL)]
        private IWebElement _suppliersSelectAll;
        [FindsBy(How = How.XPath, Using = SEARCH_SUPPLIER)]
        private IWebElement _searchSupplier;
        [FindsBy(How = How.XPath, Using = SUPPLIER_UNSELECT_ALL)]
        private IWebElement _suppliersUnselectAll;
        [FindsBy(How = How.XPath, Using = COUNT_SELECTED_MENUS)]
        private IWebElement _count_selectedmenus;
        [FindsBy(How = How.XPath, Using = SHOW_INACTIVE_SITES)]
        private IWebElement _showInactiveSites;
        [FindsBy(How = How.XPath, Using = FILTRED_SITES_LIST)]
        private IWebElement _filtredSitesList;

        public enum FilterType
        {
            SearchByName,
            Site,
            Supplier,
            ShowInactiveSuppliers,
            SupplierMultiple,
            ShowInactiveSites
        }

        public void Filter(FilterType filterType, object value, bool clearAllPreviousSelection = true, bool ignoreAutoSuggest = false)
        {
            Actions action = new Actions(_webDriver);
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            switch (filterType)
            {
                case FilterType.SearchByName:
                    _searchByNameFilter = WaitForElementIsVisibleNew(By.Id(SEARCH_BY_NAME));
                    _searchByNameFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    WaitPageLoading();
                    break;

                case FilterType.Supplier:
                    ComboBoxSelectById(new ComboBoxOptions(SUPPLIER, (string)value, false));
                    _supplier.Click();
                    break;
                case FilterType.ShowInactiveSuppliers:
                    _showinactivesuppliers = WaitForElementIsVisibleNew(By.XPath(SHOW_INACTIVE_SUPPLIERS));
                    _showinactivesuppliers.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.Site:
                    ComboBoxSelectById(new ComboBoxOptions(SITE_FILTER, (string)value, false));
                    _siteFilter.Click();
                    break;

                case FilterType.SupplierMultiple:
                    ComboBoxSelectById(new ComboBoxOptions(SUPPLIER, (string)value, false));
                    break;
                case FilterType.ShowInactiveSites:
                    _showInactiveSites = WaitForElementIsVisibleNew(By.XPath(SHOW_INACTIVE_SITES));
                    _showInactiveSites.SetValue(ControlType.CheckBox, value);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void Fill(string supplier, string group = null)
        {
            // combo box supplier (All)
            WaitPageLoading();
            _supplier = WaitForElementIsVisible(By.Id(SUPPLIER));
            if (supplier == null)
            {
                _supplier.Click();

                _selectAllSuppliers = WaitForElementIsVisible(By.XPath(SELECT_ALL_SUPPLIERS));
                _selectAllSuppliers.Click();
                // on referme
                _supplier.Click();
            }
            else
            {
                ComboBoxSelectById(new ComboBoxOptions("SelectedSupplierIds_ms", supplier, false));
                _supplier.Click();
            }

            WaitPageLoading();
        }

        public void ClickOnShowInactiveSiteCheckbox()
        {
            _showInactiveSiteCheckbox = WaitForElementIsVisible(By.XPath(SHOW_INACTIVE_SITE_CHECKBOX));
            _showInactiveSiteCheckbox.Click();
            WaitForLoad();
        }
        public void SelectMultipleInactiveSites(string firstSite, string secondSite)
        {
            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", firstSite, false));
            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", secondSite, false) { ClickUncheckAllAtStart = false });
            //_sitesCombobox = WaitForElementIsVisible(By.Id(SITES_COMBOBOX));
            //_sitesCombobox.Click();
            //WaitForLoad();
            //_unselectAllSites = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_SITES));
            //_unselectAllSites.Click();
            //WaitForLoad();
            //_siteInputCombobox = WaitForElementIsVisible(By.XPath(SITE_INPUT_COMBOBOX));
            //_siteInputCombobox.SetValue(ControlType.TextBox, "Inactive");
            //WaitForLoad();
            //_firstSiteCheckbox = WaitForElementIsVisible(By.XPath(FIRST_ITEM_SITE_COMBOBOX));
            //_firstSiteCheckbox.Click();
            //_secondSiteCheckbox = WaitForElementIsVisible(By.XPath(SECOND_ITEM_SITE_COMBOBOX));
            //_secondSiteCheckbox.Click();
        }
        public void SelectInactiveSites(string firstSite)
        {
            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", firstSite, false));
         
        }
        public bool CheckInactiveSites()
        {
            _sitesCombobox = WaitForElementIsVisible(By.Id(SITES_COMBOBOX));
            _sitesCombobox.Click();
            WaitPageLoading();
            WaitForLoad();

            _siteInputCombobox = WaitForElementIsVisible(By.XPath(SITE_INPUT_COMBO));
            WaitForLoad();
            _siteInputCombobox.SetValue(ControlType.TextBox, "Inactive");
            WaitForLoad();
            _firstSiteCheckbox = WaitForElementIsVisible(By.XPath(FIRST_ITEM_SITE_COMBO));

            return _firstSiteCheckbox.Text.Contains("Inactive");
        }
        public void SelectFirstItem(string site)
        {
            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", site, false));
            //_sitesCombobox = WaitForElementIsVisible(By.Id(SITES_COMBOBOX));
            //_sitesCombobox.Click();
            //WaitForLoad();
            //_firstSiteCheckbox = WaitForElementIsVisible(By.XPath(FIRST_ITEM_SITE_COMBOBOX));
            //_firstSiteCheckbox.Click();
            //WaitForLoad();
        }
        public void ClickSearch()
        {
            //SEARCH
            var searchAction = WaitForElementIsVisibleNew(By.Id(SEARCH_BTN));
            searchAction.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void ClickSelectAllButton()
        {
            var btn = WaitForElementIsVisibleNew(By.Id(MASSIVE_DELETE_SELECTALL_BUTTON));
            btn.Click();
            WaitForLoad();
        }    
        public void ClickUnselectAllButton()
        {
            var btn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_UNSELECTALL_BUTTON));
            btn.Click();
            WaitForLoad();
        }
        public List<string> GetSuppliersListFiltred()
        {
            var listElement = _webDriver.FindElement(By.XPath("/html/body/div[17]/ul"));
            var ordersText = listElement.Text.Trim();
            var ordersList = ordersText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(order => order.Trim()).ToList();
            return ordersList;

        }

        public int CheckTotalSelectCount()
        {
            var count = _webDriver.FindElement(By.Id(MASSIVE_DELETE_ITEM_SELECT_COUNT));
            var totalnumber = count.Text.Trim();
            return Convert.ToInt32(totalnumber);
        }
        public bool IsPageSizeEqualsTo(string size)
        {
            var nbPages = WaitForElementExists(By.XPath(DDL_PAGESIZE));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");
            return selectedValue == size;
        }
        public int GetTotalRowsForPagination()
        {
            var table = _webDriver.FindElement(By.Id("tableItems"));

            var rows = table.FindElements(By.XPath("//*[@id=\"tableItems\"]/tbody/tr[*]"));

            var totalRows = rows.Count();
            return totalRows;
        }
        public new void PageSize(string size)
        {
            string pagesizetionxpath = MASSIVE_DELETE_PAGESIZE;
            if (!isElementExists(By.XPath(pagesizetionxpath)) && size == "30")
            {
                pagesizetionxpath = "/html/body/div[3]/div/div/div[2]/div/form/div/div[6]/div/div/div/nav/select";
            }
            if (!isElementExists(By.XPath(pagesizetionxpath)) && size == "50")
            {
                pagesizetionxpath = "/html/body/div[3]/div/div/div[2]/div/form/div/div[6]/div/div/div/nav/select";
            }
            if (!isElementExists(By.XPath(pagesizetionxpath)) && size == "100")
            {
                pagesizetionxpath = "/html/body/div[3]/div/div/div[2]/div/form/div/div[6]/div/div/div/nav/select";
            }
            var PageSizeDdl = _webDriver.FindElement(By.XPath(pagesizetionxpath));
            PageSizeDdl.SetValue(ControlType.DropDownList, size);
            WaitForLoad();
        }
        public void DeleteFirstService()
        {
            SelectFirst();
            if (CheckTotalSelectCount() == 1)
            {
                var DeleteBtn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_DELETE_BTN));
                DeleteBtn.Click();
                var confirmBtn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_CONFIRM_BTN));
                confirmBtn.Click();
                var OkBtn = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SERVICE_DELETEOK_BTN));
                OkBtn.Click();
                WaitForLoad();
            }
        }
        public void SelectFirst()
        {
            string checkBoxId = MASSIVE_DELETE_SELECT_CHECKBOX;
            var checkbox = _webDriver.FindElements(By.Id(checkBoxId)).FirstOrDefault();
            checkbox.SetValue(ControlType.CheckBox, true);
            WaitForLoad();

        }

        public void SetPageSize(string size)
        {
            try
            {
                WaitForElementIsVisible(By.XPath(PAGESIZE));
                _pageSize = _webDriver.FindElement(By.XPath(PAGESIZE));
            }
            catch
            {
                // tableau vide : pas de PageSize
                return;
            }
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_pageSize).Perform();
            _pageSize.SetValue(ControlType.DropDownList, size);

            WaitPageLoading();
        }

        public List<string> GetAllSites()
        {
            List<string> allSitesResult = new List<string>();

            var allSites = _webDriver.FindElements(By.XPath(ALL_SITES));

            if (allSites != null)
            {
                foreach (var site in allSites)
                {
                    allSitesResult.Add(site.Text.Trim());
                }
            }
            return allSitesResult;
        }

        public List<string> GetPurchasingItemSite()
        {
            List<string> liste = new List<string>();
            var stars = _webDriver.FindElements(By.XPath("//*[@id=\"tableItems\"]/tbody/tr[*]/td[4]"));
            foreach (var star in stars)
            {
                liste.Add(star.Text);
            }
            return liste;
        }

        public List<string> GetPurchasingItemStatus()
        {
            List<string> liste = new List<string>();
            var stars = _webDriver.FindElements(By.XPath("//*[@id=\"tableItems\"]/tbody/tr[*]/td[3]"));
            foreach (var star in stars)
            {
                liste.Add(star.Text);
            }
            return liste;
        }

        public void ClickSite()
        {
            _selectSite = WaitForElementIsVisible(By.XPath(CLICK_SITE));
            _selectSite.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public IEnumerable<string> getItems()
        {
            return _webDriver.FindElements(By.XPath(MASSIVE_ITEMS)).Select(e => e.Text.Trim());
        }
        public IEnumerable<string> getListeItems()
        {
            return _webDriver.FindElements(By.XPath(MASSIVE_ITEMS)).Select(e => e.Text.Trim());
        }


        public void SortByItemBtn()
        {
            var SortByItemBtn = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SORT_BY_ITEMS));
            SortByItemBtn.Click();
            WaitForLoad();
        }
        public bool IsSortedAlphabetically(bool ascending = true)
        {
            var items = getItems().ToList();

            for (var i = 0; i < items.Count - 1; i++)
            {
                int comparisonResult = string.Compare(items[i], items[i + 1], StringComparison.OrdinalIgnoreCase);

                if (ascending)
                {
                    if (comparisonResult > 0) // If current item is greater than the next item, it's not sorted in ascending order
                    {
                        Console.WriteLine($"Sort check failed (ascending): '{items[i]}' should come before '{items[i + 1]}'");
                        return false;
                    }
                }
                else
                {
                    if (comparisonResult < 0) // If current item is less than the next item, it's not sorted in descending order
                    {
                        Console.WriteLine($"Sort check failed (descending): '{items[i]}' should come after '{items[i + 1]}'");
                        return false;
                    }
                }
            }

            return true;
        }


        public void ClickStatus()
        {
            _selectStatus = WaitForElementIsVisible(By.XPath(CLICK_STATUS));
            _selectStatus.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void Delete()
        {
            _deleteButton = WaitForElementIsVisibleNew(By.Id(DELETE_BUTTON));
            _deleteButton.Click();
            WaitForLoad();
            _confirmDelete = WaitForElementIsVisibleNew(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitLoading();
            Thread.Sleep(5000);
            _closePopup = WaitForElementIsVisibleNew(By.XPath(CLOSE_POPUP));
            _closePopup.Click();
            WaitForLoad();
            WaitPageLoading();
        }
        public bool VerifyRowsAreSelected()
        {
            string checkBoxXpath = "//*[@id=\"item_IsSelected\"]";
            this.SetPageSize("100");
            ClickSelectAllButton();
            WaitForLoad();
            if (this.VerifyPageRowsAreSelected(checkBoxXpath) == false)
            {
                return false;
            }
            this.GoToNextPage("2");
            WaitForLoad();
            return this.VerifyPageRowsAreSelected(checkBoxXpath);
        }

        private bool VerifyPageRowsAreSelected(string checkBoxXpath)
        {
            var rowsChecks = _webDriver.FindElements(By.XPath(checkBoxXpath));
            WaitForLoad();
            foreach (var row in rowsChecks)
            {
                try
                {
                    if (row.Selected == false)
                    {
                        return false;
                    }

                }
                catch
                {
                    return false;
                }
            }
            return true;
        }
        public void GoToNextPage(string pageNumber)
        {
            var pageNumbers = _webDriver.FindElements(By.XPath("//*[@id=\"list-items-deletion\"]/nav/ul/li[*]/a"));
            foreach (var page in pageNumbers)
            {
                if (page.Text == pageNumber)
                {
                    string xpath = "//*[@id=\"list-items-deletion\"]/nav/ul/li[*]/a[text()=" + pageNumber + "]";
                    var pageBtn = WaitForElementIsVisible(By.XPath(xpath));
                    pageBtn.Click();
                }
            }
        }
        public void SelectAllSites()
        {
            var searchSite = WaitForElementIsVisible(By.Id(SITE_FILTER));
            searchSite.Click();

             var unselectAllSite = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_SITES));
            unselectAllSite.Click();

            var selectAllSite = WaitForElementIsVisible(By.XPath(SITES_SELECT_ALL));
            selectAllSite.Click();
            searchSite.Click();

        }
        public void SelectAllSupplier()
        {
            var searchSuppiler = WaitForElementIsVisible(By.Id(SUPPLIER));
            searchSuppiler.Click();

            var unselectAllSupplier = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_SUPPLIERS));
            unselectAllSupplier.Click();

            var selectAllSupplier = WaitForElementIsVisible(By.XPath(SELECT_ALL_SUPPLIERS));
            selectAllSupplier.Click();
            searchSuppiler.Click();
        }
        public string GetFirstItem()
        {
            if (isElementVisible(By.XPath("//*[@id=\"tableItems\"]/tbody/tr[1]/td[2]")))
            {
                var FirstItem = WaitForElementIsVisible(By.XPath("//*[@id=\"tableItems\"]/tbody/tr[1]/td[2]"));
                WaitForLoad();
                return FirstItem.Text;
            }
            else
            {
                return null;
            }

        }
        public ItemGeneralInformationPage ClickOnItemLinkFromRow(int rowNumber)
        {
            var recipeLink = WaitForElementIsVisible(By.XPath("//*[@id=\"tableItems\"]/tbody/tr[" + rowNumber + "]/td[8]/a"));
            recipeLink.Click();
            WaitForLoad();
            // switch driver to the opened tab
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();
            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

  
        public void ClickSupplier()
        {
            _selectSupplier = WaitForElementIsVisible(By.XPath(CLICK_SUPPLIER));
            _selectSupplier.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public List<string> GetPurchasingItemSupplier()
        {
            List<string> liste = new List<string>();
            var stars = _webDriver.FindElements(By.XPath("//*[@id=\"tableItems\"]/tbody/tr[*]/td[5]"));
            foreach (var star in stars)
            {
                liste.Add(star.Text);
            }
            return liste;
        }

        public List<string> GetPurchasingItemIDStatus()
        {
            WaitForLoad();
            List<string> liste = new List<string>();
            var stars = _webDriver.FindElements(By.XPath("//*[@id=\"tableItems\"]/tbody/tr[*]/td[6]"));
            foreach (var star in stars)
            {
                liste.Add(star.Text);
            }
            return liste;
        }

        internal bool HasData()
        {
            return isElementVisible(By.XPath("//*[@id=\"tableItems\"]/tbody/tr/td[2]/b"));
        }
        public int CheckSelectedMenus()
        {
            WaitForLoad();
            _count_selectedmenus = WaitForElementExists(By.Id(COUNT_SELECTED_MENUS));
            int nombre = Int32.Parse(_count_selectedmenus.Text);
            return nombre;
        }
        public List<string> GetSitesListFiltred()
        {
             _filtredSitesList = _webDriver.FindElement(By.XPath(FILTRED_SITES_LIST));
            var ordersText = _filtredSitesList.Text.Trim();
            var ordersList = ordersText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(order => order.Trim()).ToList();
            return ordersList;

        }
        public string GetSupplierName()
        {
            var name = WaitForElementIsVisible(By.XPath(SUPLLIER_NAME));
            return name?.Text; 
        }
        public void UncheckAllSites()
        {
            IWebElement input = WaitForElementIsVisible(By.Id(SITE_FILTER));
            WaitForLoad();
            input.Click();
            WaitForLoad();

            if (isElementVisible(By.XPath(INPUT_SEARCH)))
            {
                var inputSearch = WaitForElementIsVisible(By.XPath(INPUT_SEARCH));
                inputSearch.Clear();
            }
    
              var uncheckAllVisible = SolveVisible("//*/span[text()='Uncheck all']");
            Assert.IsNotNull(uncheckAllVisible);
            uncheckAllVisible.Click();
            input.SendKeys(Keys.Enter);


        }
        public bool GetTotalTabs()
        {
            IList<string> allWindows = _webDriver.WindowHandles;

            return allWindows.Count > 1;
        }
        public string GetItemName()
        {
            if (isElementVisible(By.XPath("//*[@id=\"tableItems\"]/tbody/tr[1]/td[2]")))
            {
                var ItemName = WaitForElementIsVisible(By.XPath("//*[@id=\"tableItems\"]/tbody/tr[1]/td[2]")); 
                WaitForLoad();
                return ItemName.Text;
            }
            else
            {
                return null;
            }

        }
    }
}
