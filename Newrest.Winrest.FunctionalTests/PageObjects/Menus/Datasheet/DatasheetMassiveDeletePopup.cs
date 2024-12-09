using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet
{
    public class DatasheetMassiveDeletePopup : PageBase
    {
        public class MassiveDeleteStatus
        {
            private MassiveDeleteStatus(string value) { Value = value; }

            public string Value { get; private set; }

            public static MassiveDeleteStatus ActiveDatasheets { get { return new MassiveDeleteStatus("Active datasheet"); } }
            public static MassiveDeleteStatus InactiveDatasheets { get { return new MassiveDeleteStatus("Inactive datasheet"); } }
            public static MassiveDeleteStatus OnlyInactiveSites { get { return new MassiveDeleteStatus("Only inactive sites"); } }
            public static MassiveDeleteStatus OnlyInactiveCustomers { get { return new MassiveDeleteStatus("Only inactive customers"); } }
            public static MassiveDeleteStatus Unused { get { return new MassiveDeleteStatus("Unused"); } }
            public static MassiveDeleteStatus Used { get { return new MassiveDeleteStatus("Used"); } }
            public static MassiveDeleteStatus WithUnpurchasableItems { get { return new MassiveDeleteStatus("With unpurchasable items"); } }

            public override string ToString()
            {
                return Value;
            }
        }

        public DatasheetMassiveDeletePopup(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        #region Constantes

        private const string SEARCH_PATTERN_XPATH = "//*[@id=\"formMassiveDeleteDatasheet\"]/descendant::input[contains(@id, 'SearchPattern')]";
        private const string SITE_FILTER_ID = "siteFilter";
        private const string SITE_FILTER_SHOW_INACTIVE_ID = "ShowInactiveSites";
        private const string CUSTOMER_FILTER_ID = "customerFilter";
        private const string CUSTOMER_FILTER_SHOW_INACTIVE_ID = "ShowInactiveCustomers";
        private const string GUESTTYPE_FILTER_ID = "guestTypeFilter";
        private const string STATUS_FILTER_ID = "statusFilter";
        private const string SEARCH_BTN_ID = "SearchDatasheetsBtn";
        private const string SELECTALL_BTN_ID = "selectAll";
        private const string DELETE_BTN = "deleteDatasheetBtn";
        private const string CONFIRM_DELETE = "dataConfirmOK";
        private const string UNSELECT = "unselectAll";
        private const string SELECTALL = "selectAll";
        private const string PAGE_SIZE = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/nav/select";
        private const string ROWS_FOR_PAGINATION = "//*[@id=\"tableDatasheets\"]/div[*]/div/div/div[2]/table/tbody/tr/td[2]";
        private const string DATASHEET_BUTTON = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/thead/tr/th[2]/span/a";
        private const string CUSTOMERCODE_BTN = "//*[@id=\"tableDatasheets\"]/thead/tr/th[4]/span/a";
        private const string GUESTTYPE_BTN = "//*[@id=\"tableDatasheets\"]/thead/tr/th[5]/span/a";
        private const string USECASE_BTN = "//*[@id=\"tableDatasheets\"]/thead/tr/th[6]/span/a";
        private const string DDL_PAGESIZE = "//*[@id=\"div-datasheetResultTable\"]/descendant::select[contains(@id, 'page-size-selector')]";

        private const string DATASHEETSITE_BUTTON = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/thead/tr/th[3]/span/a";
        private const string TOTAL_USE_CASES = "//*[@id=\"tableDatasheets\"]/tbody/tr[*]/td[6]";
        private const string CHECK_ALL = "//*[@id=\"selectAll\"]";
        private const string DELETE_BUTN = "//*[@id=\"deleteDatasheetBtn\"]";
        private const string OK = "//*[@id=\"dataConfirmOK\"]";
        private const string CONFIRM = "//*[@id=\"modal-1\"]/div[3]/button";


        #endregion

        #region Composants utiles sélectionnables à tout moment

        [FindsBy(How = How.XPath, Using = SEARCH_PATTERN_XPATH)]
        private IWebElement _searchPattern;

        [FindsBy(How = How.Id, Using = SITE_FILTER_ID)]
        private IWebElement _sitesFilter;

        [FindsBy(How = How.Id, Using = CUSTOMER_FILTER_ID)]
        private IWebElement _customersFilter;

        [FindsBy(How = How.Id, Using = GUESTTYPE_FILTER_ID)]
        private IWebElement _guestTypeFilter;

        [FindsBy(How = How.Id, Using = STATUS_FILTER_ID)]
        private IWebElement _statusFilter;

        [FindsBy(How = How.XPath, Using = PAGE_SIZE)]
        private IWebElement _pageSize;

        [FindsBy(How = How.Id, Using = DATASHEET_BUTTON)]
        private IWebElement _dataSheetButton;

        [FindsBy(How = How.Id, Using = DATASHEETSITE_BUTTON)]
        private IWebElement _dataSheetSiteButton;
        #endregion

        public void SetDatasheetName(string datasheetName)
        {
            _searchPattern = WaitForElementIsVisible(By.XPath(SEARCH_PATTERN_XPATH));
            WaitForLoad();
            _searchPattern.SetValue(ControlType.TextBox, datasheetName);
        }

        public void SelectSiteByName(string siteName, bool ignoreUncheckAll)
        {
            Actions action = new Actions(_webDriver);

            _sitesFilter = WaitForElementExists(By.Id(SITE_FILTER_ID));
            action.MoveToElement(_sitesFilter).Perform();
            ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER_ID, siteName, false) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = !ignoreUncheckAll };
            ComboBoxSelectById(cbOpt);
        }

        public void SelectAllInactiveSites()
        {
            Actions action = new Actions(_webDriver);

            _sitesFilter = WaitForElementExists(By.Id(SITE_FILTER_ID));
            action.MoveToElement(_sitesFilter).Perform();
            ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER_ID, "Inactive", true)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }

        public void ClickOnInactiveSiteCheck()
        {
            IWebElement checkBoxInactiveSite = WaitForElementExists(By.Id(SITE_FILTER_SHOW_INACTIVE_ID));
            checkBoxInactiveSite.Click();
            WaitForLoad();
        }

        public void SelectCustomerByName(string customerName, bool ignoreUncheckAll)
        {
            Actions action = new Actions(_webDriver);

            _customersFilter = WaitForElementExists(By.Id(CUSTOMER_FILTER_ID));
            action.MoveToElement(_customersFilter).Perform();
            ComboBoxOptions cbOpt = new ComboBoxOptions(CUSTOMER_FILTER_ID, customerName, false) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = !ignoreUncheckAll };
            ComboBoxSelectById(cbOpt);
        }

        public void SelectAllInactiveCustomers()
        {
            Actions action = new Actions(_webDriver);

            _customersFilter = WaitForElementExists(By.Id(CUSTOMER_FILTER_ID));
            action.MoveToElement(_customersFilter).Perform();
            ComboBoxOptions cbOpt = new ComboBoxOptions(CUSTOMER_FILTER_ID, "Inactive", false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }

        public void ClickOnInactiveCustomerCheck()
        {
            IWebElement checkBoxInactiveCustomer = WaitForElementExists(By.Id("ShowInactiveCustomers"));
            checkBoxInactiveCustomer.Click();
            WaitForLoad();
        }

        public void SelectGuestTypeByName(string guestTypeName, bool ignoreUncheckAll)
        {
            Actions action = new Actions(_webDriver);

            _guestTypeFilter = WaitForElementExists(By.Id(GUESTTYPE_FILTER_ID));
            action.MoveToElement(_guestTypeFilter).Perform();
            ComboBoxOptions cbOpt = new ComboBoxOptions(GUESTTYPE_FILTER_ID, guestTypeName, false) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = !ignoreUncheckAll };
            ComboBoxSelectById(cbOpt);
        }

        /// <summary>
        /// L'index à sélectionner dans le multi select du statut. L'index commence à 1.
        /// </summary>
        /// <param name="index"></param>
        public void SelectStatus(string statusLabel)
        {
            Actions action = new Actions(_webDriver);

            _statusFilter = WaitForElementExists(By.XPath("//*[@id=\"collapseSelectedStatusFilter\"]"));
            action.MoveToElement(_statusFilter).Perform();
            ComboBoxOptions cbOpt = new ComboBoxOptions()
            {
                XpathId = "collapseSelectedStatusFilter",
                SelectionValue = statusLabel,
                ClickCheckAllAtStart = false,
                IsUsedInFilter = false
            };
            ComboBoxSelectById(cbOpt);
        }

        public void ClickOnSearch()
        {
            WaitPageLoading();
            var searchBtn = WaitForElementExists(By.Id(SEARCH_BTN_ID));
            searchBtn.Click();
            WaitPageLoading();
        }

        public void ClickDatasheetNameButton()
        {
            _dataSheetButton = WaitForElementIsVisible(By.XPath(DATASHEET_BUTTON));
            _dataSheetButton.Click();
            WaitPageLoading();
        }

        public void ClickSiteButton()
        {
            _dataSheetSiteButton = WaitForElementIsVisible(By.XPath(DATASHEETSITE_BUTTON));
            _dataSheetSiteButton.Click();
            WaitPageLoading();
        }

        public bool VerifySortByDatasheetName()
        {
            var isOrdered = true;
            this.SetPageSize("100");
            List<string> namesText = new List<string>();
            var names = _webDriver.FindElements(By.XPath("//*[@id=\"tableDatasheets\"]/tbody/tr[*]/td[2]"));
            foreach (var name in names)
            {
                namesText.Add(name.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < namesText.Count - 1; i++)
            {
                int comparisonResult = string.Compare(namesText[i], namesText[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult > 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public void ClickOnCustomerHeader()
        {
            var customerHeader = WaitForElementExists(By.XPath(CUSTOMERCODE_BTN));
            customerHeader.Click();
            WaitPageLoading();
        }

        public void ClickOnGuestTypeHeader()
        {
            var guestTypeHeader = WaitForElementExists(By.XPath(GUESTTYPE_BTN));
            guestTypeHeader.Click();
            WaitPageLoading();
        }

        public bool VerifySortByDatasheetSite()
        {
            var isOrdered = true;
            this.SetPageSize("100");
            List<string> sitesText = new List<string>();
            var siteNames = _webDriver.FindElements(By.XPath("//*[@id=\"tableDatasheets\"]/tbody/tr[*]/td[3]"));
            foreach (var name in siteNames)
            {
                sitesText.Add(name.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < sitesText.Count - 1; i++)
            {
                int comparisonResult = string.Compare(sitesText[i], sitesText[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult > 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public void ClickOnUseCaseHeader()
        {
            var usecaseHeader = WaitForElementExists(By.XPath(USECASE_BTN));
            usecaseHeader.Click();
            WaitPageLoading();
        }

        public void ClickOnSelectAll()
        {
            var selectAllBtn = WaitForElementExists(By.Id(SELECTALL));
            selectAllBtn.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
        }

        public int CheckTotalNumberDataSheet()
        {
            WaitPageLoading();
            ClickOnSelectAll();
            var _totalNumber = WaitForElementExists(By.Id("datasheetCount"));
            int nombre = Int32.Parse(_totalNumber.Text);
            return nombre;
        }

        public bool IsPageSizeEqualsTo16()
        {
            var nbPages = WaitForElementExists(By.XPath(DDL_PAGESIZE));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");

            return selectedValue == "16";
        }

        public bool IsPageSizeEqualsTo30()
        {
            var nbPages = WaitForElementExists(By.XPath(DDL_PAGESIZE));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");

            return selectedValue == "30";
        }

        public bool IsPageSizeEqualsTo50()
        {
            var nbPages = WaitForElementExists(By.XPath(DDL_PAGESIZE));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");

            return selectedValue == "50";
        }

        public new bool IsPageSizeEqualsTo100()
        {
            var nbPages = WaitForElementExists(By.XPath(DDL_PAGESIZE));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");

            return selectedValue == "100";
        }

        public int GetTotalRowsForPagination()
        {
            var table = _webDriver.FindElement(By.Id("tableDatasheets"));

            var rows = table.FindElements(By.TagName("tr"));

            // -1  Subtract 1 from the count to exclude the first row
            return rows.Count - 1;
        }

        public DatasheetMassiveDeleteRowResult GetRowResultInfo(int itemOffset)
        {
            string rowXpath = "(//*[@id=\"tableDatasheets\"]/tbody/tr)[" + (itemOffset + 1) + "]";
            IWebElement rowResult = null;
            try
            {
                rowResult = WaitForElementIsVisible(By.XPath(rowXpath));
            }
            catch
            { //silent catch : ignore if element isn't found
            }

            if (rowResult == null)
            {
                return null;
            }
            else
            {
                DatasheetMassiveDeleteRowResult dsRowResult = new DatasheetMassiveDeleteRowResult();

                dsRowResult.DatasheetName = rowResult.FindElement(By.XPath(rowXpath + "/td[2]")).Text;

                if (string.IsNullOrEmpty(dsRowResult.DatasheetName) || dsRowResult.DatasheetName.StartsWith("No data"))
                { return null; }

                dsRowResult.IsDatasheetInactive = rowResult.FindElement(By.XPath(rowXpath + "/td[2]")).GetAttribute("class").Contains("IsInactive");
                dsRowResult.SiteName = rowResult.FindElement(By.XPath(rowXpath + "/td[3]")).Text;
                dsRowResult.IsSiteInactive = rowResult.FindElement(By.XPath(rowXpath + "/td[3]")).GetAttribute("class").Contains("IsInactive");
                dsRowResult.CustomerName = rowResult.FindElement(By.XPath(rowXpath + "/td[4]")).Text;
                dsRowResult.IsCustomerInactive = rowResult.FindElement(By.XPath(rowXpath + "/td[4]")).GetAttribute("class").Contains("IsInactive");
                dsRowResult.GuestTypeName = rowResult.FindElement(By.XPath(rowXpath + "/td[5]")).Text;
                dsRowResult.UseCase = int.Parse(rowResult.FindElement(By.XPath(rowXpath + "/td[6]")).Text);
                return dsRowResult;
            }
        }

        public void GoToNextPage(string pageNumber)
        {
            var pageNumbers = _webDriver.FindElements(By.XPath("//*[@id=\"list-datasheets-deletion\"]/nav/ul/li[*]/a"));
            foreach (var page in pageNumbers)
            {
                if (page.Text == pageNumber)
                {
                    string xpath = "//*[@id=\"list-datasheets-deletion\"]/nav/ul/li[*]/a[text()=" + pageNumber + "]";
                    var pageBtn = WaitForElementIsVisible(By.XPath(xpath));
                    pageBtn.Click();
                }
            }
        }

        private List<string> GetCustomersCodes()
        {
            List<string> ids = new List<string>();
            string xPath = "//*[@id=\"tableDatasheets\"]/tbody/tr[*]/td[4]";
            var customerCodes = _webDriver.FindElements(By.XPath(xPath));

            foreach (var code in customerCodes)
            {
                ids.Add(code.Text);
            }
            return ids;
        }

        private List<string> GetGuestTypes()
        {
            List<string> ids = new List<string>();
            string xPath = "//*[@id=\"tableDatasheets\"]/tbody/tr[*]/td[5]";
            var guestTypes = _webDriver.FindElements(By.XPath(xPath));

            foreach (var guestType in guestTypes)
            {
                ids.Add(guestType.Text);
            }
            return ids;
        }

        private List<int> GetUseCase()
        {
            List<int> usecases = new List<int>();
            string xPath = "//*[@id=\"tableDatasheets\"]/tbody/tr[*]/td[6]";
            var usecaseResults = _webDriver.FindElements(By.XPath(xPath));

            foreach (var uc in usecaseResults)
            {
                usecases.Add(int.Parse(uc.Text));
            }
            return usecases;
        }

        public void SetPageSize(string size)
        {
            try
            {
                WaitForElementIsVisible(By.XPath(DDL_PAGESIZE));
                _pageSize = _webDriver.FindElement(By.XPath(DDL_PAGESIZE));
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

        public bool VerifyCustomerCodeSort(SortType sortType)
        {
            this.SetPageSize("100");

            List<string> ids = this.GetCustomersCodes();

            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if ((sortType == SortType.Ascending && comparisonResult > 0) || (sortType == SortType.Descending && comparisonResult < 0))
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifyGuestTypeSort(SortType sortType)
        {
            this.SetPageSize("100");

            List<string> ids = this.GetGuestTypes();

            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if ((sortType == SortType.Ascending && comparisonResult > 0) || (sortType == SortType.Descending && comparisonResult < 0))
                {
                    return false;
                }
            }
            return true;
        }

        private bool VerifyPageRowsAreSelected(string checkBoxXpath)
        {
            var rowsChecks = _webDriver.FindElements(By.XPath(checkBoxXpath));
            foreach (var row in rowsChecks)
            {
                if (row.Selected == false)
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifyRowsAreSelected()
        {
            string checkBoxXpath = "//*[@id=\"item_IsSelected\"]";
            this.SetPageSize("100");
            if (this.VerifyPageRowsAreSelected(checkBoxXpath) == false)
            {
                return false;
            }
            this.GoToNextPage("2");
            return this.VerifyPageRowsAreSelected(checkBoxXpath);
        }

        public bool VerifyUseCaseSort(SortType sortType)
        {
            this.SetPageSize("100");

            List<int> usecases = this.GetUseCase();

            for (int i = 0; i < usecases.Count - 1; i++)
            {
                if ((sortType == SortType.Ascending && usecases[i] > usecases[i + 1])
                    || (sortType == SortType.Descending && usecases[i] < usecases[i + 1]))
                {
                    return false;
                }
            }
            return true;
        }

        public int GetUseCases()
        {
            WaitPageLoading();
            WaitForLoad();
            var elements = _webDriver.FindElements(By.XPath(TOTAL_USE_CASES));
            int totalUseCase = 0;
            foreach (var elm in elements)
            {
                var useCase = Int32.Parse(elm.Text.Trim());
                totalUseCase += useCase;
            }
            return totalUseCase;
        }

        public DatasheetDetailsPage GoToDatasheetPage()
        {
            var datasheetLink = WaitForElementIsVisible(By.XPath("//*[@id=\"tableDatasheets\"]/tbody/tr[1]/td[7]/a"));
            datasheetLink.Click();
            WaitForLoad();
            // switch driver to the opened tab
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();
            return new DatasheetDetailsPage(_webDriver, _testContext);
        }

        public bool ClickOnResultLinkAndCheckDatasheetName(int rowOffset)
        {
            string rowXpath = "(//*[@id=\"tableDatasheets\"]/tbody/tr)[" + (rowOffset + 1) + "]";
            IWebElement rowResult = null;
            try
            {
                rowResult = WaitForElementIsVisible(By.XPath(rowXpath));
            }
            catch
            { //silent catch : ignore if element isn't found
            }

            if (rowResult == null)
            {
                return false;
            }
            else
            {
                string datasheetName = rowResult.FindElement(By.XPath(rowXpath + "/td[2]")).Text;
                var datasheetLink = rowResult.FindElement(By.XPath(rowXpath + "/td[7]"));
                datasheetLink.Click();

                var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
                wait.Until((driver) => driver.WindowHandles.Count > 1);
                _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

                return WaitForElementExists(By.XPath("//*[@class=\"title-bar\"]/h1")).Text.Contains(datasheetName);
            }
        }
        public void SelectAllSites()
        {
            var checkAllSitesCheckbox = _webDriver.FindElement(By.XPath(CHECK_ALL));
            if (!checkAllSitesCheckbox.Selected)
            {
                checkAllSitesCheckbox.Click(); 
            }
        }
        public void ClickDeleteButton()
        {
            WaitPageLoading();
            var DeleteBtn = _webDriver.FindElement(By.XPath(DELETE_BUTN));
            DeleteBtn.Click();
            WaitPageLoading();
            _webDriver.FindElement(By.XPath(OK)).Click();
            WaitPageLoading();
            _webDriver.FindElement(By.XPath(CONFIRM)).Click();
        }


    }

    public class DatasheetMassiveDeleteRowResult
    {
        public string DatasheetName { get; set; }
        public bool IsDatasheetInactive { get; set; }
        public string SiteName { get; set; }
        public bool IsSiteInactive { get; set; }
        public string CustomerName { get; set; }
        public bool IsCustomerInactive { get; set; }
        public string GuestTypeName { get; set; }
        public int UseCase { get; set; }
    }

    public enum SortType
    {
        Ascending,
        Descending,
    }   

}
