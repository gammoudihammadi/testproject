using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery
{
    public class DeliveryMassiveDeletePage : PageBase
    {
        public class MassiveDeleteStatus
        {
            private MassiveDeleteStatus(string value) { Value = value; }

            public string Value { get; private set; }

            public static MassiveDeleteStatus ActiveDelivery { get { return new MassiveDeleteStatus("Active delivery"); } }
            public static MassiveDeleteStatus InactiveDelivery { get { return new MassiveDeleteStatus("Inactive delivery"); } }
            public static MassiveDeleteStatus OnlyInactiveSites { get { return new MassiveDeleteStatus("Only inactive sites"); } }
            public static MassiveDeleteStatus OnlyInactiveCustomers { get { return new MassiveDeleteStatus("Only inactive customers"); } }
            public static MassiveDeleteStatus NoService { get { return new MassiveDeleteStatus("No service"); } }
            public static MassiveDeleteStatus NoValidationDate { get { return new MassiveDeleteStatus("No validation date"); } }

            public override string ToString()
            {
                return Value;
            }
        }

        public DeliveryMassiveDeletePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //___________________________________ Constantes _______________________________________________
        private const string CUSTOMER_FILTER = "SelectedCustomersForDeletion_ms";
        private const string CUSTOMER_CODE_BUTTON = "//*[@data-val=\"CustomerCode\"]";
        private const string SEARCH_BUTTON = "SearchDeliveriesBtn";
        private const string CUSTOMER_CELLS = "#tableDeliveries tbody td:nth-child(5) div";
        private const string SHOW_INACTIVE_CUSTOMERS = "//*[@id=\"formMassiveDeleteDelivery\"]/div/div[3]/div[2]/div";
        private const string INACTIVE_CUSTOMERS = "//*[@id=\"tableDeliveries\"]/tbody/tr[*]/td[5]";
        private const string SEARCH_RESULT_SITE = "//*[@id=\"tableDeliveries\"]/tbody/tr/td[4]/div";
        private const string SITE_FILTER = "SelectedSiteIds_ms";
        private const string SEARCH_PATTERN = "//*[@id=\"formMassiveDeleteDelivery\"]/div/div[1]/div/input[@id=\"SearchPattern\"]";
        private const string DELIVERY_BUTTON = "//*[@id=\"tableDeliveries\"]/thead/tr/th[2]/span/a";
        private const string PAGE_SIZE = "/html/body/div[3]/div/div/div[2]/div/form/div/div[7]/div/div/nav/select";
        private const string ROWS_FOR_PAGINATION = "//*[@id=\"tableDatasheets\"]/div[*]/div/div/div[2]/table/tbody/tr/td[2]";
        private const string SITE_BUTTON = "//*[@id=\"tableDeliveries\"]/thead/tr/th[2]/span/a";
        private const string SERVICE_BUTTON = "//*[@id=\"tableDeliveries\"]/thead/tr/th[3]/span/a";
        private const string DELROUND_BUTTON = "//*[@id=\"tableDeliveries\"]/thead/tr/th[6]/span/a";
        private const string VALIDATION_DATE_BUTTON = "//*[@id=\"tableDeliveries\"]/thead/tr/th[7]/span/a";

        private const string SELECT_ALL = "selectAll";
        private const string STATUS_FILTER = "SelectedStatus_ms";
        private const string INACTIVE_SITE_CHECKBOX = "//*/label[@for='ShowInactiveSites']/../div/input";
        private const string SITE_SELECT_SEARCH = "/html/body/div[16]/ul/li[not(@class='ui-multiselect-excluded')]/label";
        private const string MASSIVE_DELETE_PAGINATION = "//*[@id='list-deliveries-deletion']/nav/ul/li[position()=last()-2]/a";
        private const string MASSIVE_DELETE_SERVICE_SITE_LIST = "//*[@id=\"tableDeliveries\"]/tbody/tr[*]/td[4]";
        private const string DELETEDELIVERIESBTN = "deleteDeliveriesBtn";
        private const string MASSIVE_DELETE_ITEM_SELECT_COUNT = "deliveryCount";

        private const string CONFIRMATION_DELETE = "dataConfirmOK";
        private const string CONFIRMATIONMESSAGE = "//*[@id=\"modal-1\"]/div[2]/div/h4";


        //____________________________________ Variables _______________________________________________
        [FindsBy(How = How.Id, Using = CONFIRMATION_DELETE)]
        private IWebElement _confirmationdelete;

        [FindsBy(How = How.Id, Using = DELETEDELIVERIESBTN)]
        private IWebElement _deleteDeliveriesBtn;

        [FindsBy(How = How.Id, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.Id, Using = CUSTOMER_CODE_BUTTON)]
        private IWebElement _custormerCodeButton;

        [FindsBy(How = How.Id, Using = VALIDATION_DATE_BUTTON)]
        private IWebElement _deliveryButton;

        [FindsBy(How = How.Id, Using = SERVICE_BUTTON)]
        private IWebElement _serviceButton;

        [FindsBy(How = How.Id, Using = DELROUND_BUTTON)]
        private IWebElement _delroundButton;

        [FindsBy(How = How.Id, Using = DELROUND_BUTTON)]
        private IWebElement _valDateButton;

        [FindsBy(How = How.Id, Using = SITE_BUTTON)]
        private IWebElement _siteButton;

        [FindsBy(How = How.Id, Using = SEARCH_BUTTON)]
        private IWebElement _searchButton;

        [FindsBy(How = How.CssSelector, Using = CUSTOMER_CELLS)]
        private IList<IWebElement> _customerCells;


        [FindsBy(How = How.XPath, Using = SHOW_INACTIVE_CUSTOMERS)]
        private IWebElement _showInactiveCustomers;

        [FindsBy(How = How.XPath, Using = INACTIVE_CUSTOMERS)]
        private IList<IWebElement> _InactiveCustomers;

        [FindsBy(How = How.XPath, Using = SEARCH_RESULT_SITE)]
        private IList<IWebElement> _searchResultSites;

        [FindsBy(How = How.XPath, Using = SEARCH_PATTERN)]
        private IWebElement _searchPattern;

        [FindsBy(How = How.Id, Using = SELECT_ALL)]
        private IWebElement _selectAll;

        [FindsBy(How = How.XPath, Using = PAGE_SIZE)]
        private IWebElement _pageSize;

        public void SelectCustomer(string customer)
        {
            ComboBoxSelectById(new ComboBoxOptions(CUSTOMER_FILTER, customer, false));
            WaitForLoad();
        }

        public void SelectAllInactiveCustomers()
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(CUSTOMER_FILTER, "Inactive", false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }

        public void ClickSearchButton()
        {
            _searchButton = WaitForElementIsVisible(By.Id(SEARCH_BUTTON));
            _searchButton.Click();
            WaitPageLoading();
        }

        public void ClickDeliveryButton()
        {
            _deliveryButton = WaitForElementIsVisible(By.XPath(DELIVERY_BUTTON));
            _deliveryButton.Click();
            WaitPageLoading();
        }

        public void ClickSiteButton()
        {
            _siteButton = WaitForElementIsVisible(By.XPath(SITE_BUTTON));
            _siteButton.Click();
            WaitPageLoading();
        }
        public void ClickDelRoundButton()
        {
            _delroundButton = WaitForElementIsVisible(By.XPath(DELROUND_BUTTON));
            _delroundButton.Click();
            WaitPageLoading();
        }
        public void ClickValitionDateButton()
        {
            _valDateButton = WaitForElementIsVisible(By.XPath(VALIDATION_DATE_BUTTON));
            _valDateButton.Click();
            WaitPageLoading();
        }
        public void ClickServiceButton()
        {
            _serviceButton = WaitForElementIsVisible(By.XPath(SERVICE_BUTTON));
            _serviceButton.Click();
            WaitPageLoading();
        }

        public bool VerifyCustomerFilter(string customer)
        {
            string pattern = @"\d+";
            Match match = Regex.Match(customer, pattern);
            _customerCells = _webDriver.FindElements(By.CssSelector(CUSTOMER_CELLS));
            return _customerCells.All(cell => cell.Text == match.Value);
        }

        public void ClickCustomerCodeButton()
        {
            _custormerCodeButton = WaitForElementIsVisible(By.XPath(CUSTOMER_CODE_BUTTON));
            _custormerCodeButton.Click();
            WaitPageLoading();
        }

        public bool VerifyCustomerCodeSortById()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var customerCodes = _webDriver.FindElements(By.XPath(CUSTOMER_CODE_BUTTON));
            foreach (var code in customerCodes)
            {
                ids.Add(code.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult > 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public bool VerifyCustomerCodeSortByIdDescending()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var customerCodes = _webDriver.FindElements(By.XPath(CUSTOMER_CODE_BUTTON));
            foreach (var code in customerCodes)
            {
                ids.Add(code.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult < 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public bool VerifySiteSortByIdDescending()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var deliverySites = _webDriver.FindElements(By.XPath(SITE_BUTTON));
            foreach (var site in deliverySites)
            {
                ids.Add(site.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult < 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }


        public bool VerifySiteSortById()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var deliverySites = _webDriver.FindElements(By.XPath(SITE_BUTTON));
            foreach (var site in deliverySites)
            {
                ids.Add(site.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult > 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public bool VerifyDeliverySortById()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var deliveryNames = _webDriver.FindElements(By.XPath(DELIVERY_BUTTON));
            foreach (var name in deliveryNames)
            {
                ids.Add(name.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult > 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }
        public bool VerifyDeliverySortByIdDescending()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var deliveryNames = _webDriver.FindElements(By.XPath(DELIVERY_BUTTON));
            foreach (var name in deliveryNames)
            {
                ids.Add(name.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult < 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public bool VerifyDelRoundSortById()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var deliveryRounds = _webDriver.FindElements(By.XPath(DELROUND_BUTTON));
            foreach (var round in deliveryRounds)
            {
                ids.Add(round.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult > 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public bool VerifyDelRoundSortByIdDescending()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var deliveryRounds = _webDriver.FindElements(By.XPath(DELROUND_BUTTON));
            foreach (var round in deliveryRounds)
            {
                ids.Add(round.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult < 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public bool VerifyValitionDateSortById()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var validationDates = _webDriver.FindElements(By.XPath(DELROUND_BUTTON));
            foreach (var date in validationDates)
            {
                ids.Add(date.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult > 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public bool VerifyValitionDateSortByIdDescending()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var validationDates = _webDriver.FindElements(By.XPath(DELROUND_BUTTON));
            foreach (var date in validationDates)
            {
                ids.Add(date.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult < 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public bool VerifyServiceSortById()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var deliveryServices = _webDriver.FindElements(By.XPath(SERVICE_BUTTON));
            foreach (var service in deliveryServices)
            {
                ids.Add(service.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult > 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public bool VerifyServiceSortByIdDescending()
        {
            var isOrdered = true;
            PageSize("100");
            List<string> ids = new List<string>();
            var deliveryServices = _webDriver.FindElements(By.XPath(SERVICE_BUTTON));
            foreach (var service in deliveryServices)
            {
                ids.Add(service.Text);
            }

            // Comparing IDs to check if they are sorted
            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult < 0)
                {
                    isOrdered = false;
                }
            }
            return isOrdered;
        }

        public void SelectInactiveCustomer()
        {
            _showInactiveCustomers = WaitForElementIsVisible(By.XPath(SHOW_INACTIVE_CUSTOMERS));
            _showInactiveCustomers.Click();
            WaitForLoad();
        }

        public void SelectSites(string site)
        {
            ComboBoxSelectById((new ComboBoxOptions(SITE_FILTER, site, false)));
            WaitForLoad();
        }

        public bool VerifySearchSite(string site)
        {
            _searchResultSites = _webDriver.FindElements(By.XPath(SEARCH_RESULT_SITE));

            foreach (var sites in _searchResultSites)
            {
                string SiteResults = sites.Text;
                if (SiteResults == null && SiteResults != site)
                {
                    return false;
                }
            }

            return true;
        }

        public void SearchDeliveryName(string DeliveryName)
        {
            _searchPattern = WaitForElementIsVisible(By.XPath(SEARCH_PATTERN));
            _searchPattern.SendKeys(DeliveryName);
        }
        public bool VerifyDeliveryName(string DeliveryName)
        {
            var DeliveryResultName = WaitForElementIsVisible(By.XPath("//*[@id=\"tableDeliveries\"]/tbody/tr[1]/td[2]/div"));
            if (DeliveryResultName.Text == DeliveryName)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public DeliveryLoadingPage ClickDeliveryDetail()
        {
            var deliveryDetail = WaitForElementIsVisible(By.XPath("//*[@id=\"tableDeliveries\"]/tbody/tr[1]/td[8]/a"));
            deliveryDetail.Click();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new DeliveryLoadingPage(_webDriver, _testContext);
        }

        public string GetFirstDeliveryName()
        {
            if (isElementVisible(By.XPath("//*[@id=\"tableDeliveries\"]/tbody/tr[1]/td[2]/div")))
            {
                var DeliveryName = WaitForElementIsVisible(By.XPath("//*[@id=\"tableDeliveries\"]/tbody/tr[1]/td[2]/div"));
                return DeliveryName.Text;
            }
            else
            {
                return "";
            }

        }

        public void SelectAllCustomer()
        {
            IWebElement input = WaitForElementIsVisible(By.Id(CUSTOMER_FILTER));
            input.Click();
            var checkAllVisible = SolveVisible("//*/span[text()='Check all']");
            Assert.IsNotNull(checkAllVisible);
            checkAllVisible.Click();
        }
        public void SelectAllSites()
        {
            IWebElement input = WaitForElementIsVisible(By.Id(SITE_FILTER));
            input.Click();
            var checkAllVisible = SolveVisible("//*/span[text()='Check all']");
            Assert.IsNotNull(checkAllVisible);
            checkAllVisible.Click();
        }
        public void SelectAll()
        {
            _selectAll = WaitForElementIsVisible(By.Id(SELECT_ALL));
            _selectAll.Click();
        }
        public string GetSelectedDelivery()
        {
            var selected = WaitForElementIsVisible(By.Id("deliveryCount"));
            return selected.Text;
        }

        public int CheckTotalNumberDataSheet()
        {

            var _totalNumber = WaitForElementExists(By.Id("deliveryCount"));
            int nombre = Int32.Parse(_totalNumber.Text);
            return nombre;
        }

        public void PageSizeMassiveDeletePopup(string size)
        {
            if (size == "1000")
            {   // Test
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                js.ExecuteScript("$('#" + PAGE_SIZE + "').append($('<option>', {value: 1000,text: '1000'}),'');");
            }

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            try
            {
                if (size == "16")
                {
                    WaitForElementIsVisible(By.XPath(PAGE_SIZE));
                    _pageSize = _webDriver.FindElement(By.XPath(PAGE_SIZE));
                }
                else
                {
                    WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div/div[7]/div/div/div/nav/select"));
                    _pageSize = _webDriver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div/div[7]/div/div/div/nav/select"));
                }
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

        public bool isPageSizeEqualsTo16()
        {
            var nbPages = WaitForElementExists(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div/div[7]/div/div/div/nav/select"));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");

            return selectedValue == "16";
        }

        public bool isPageSizeEqualsTo30()
        {
            var nbPages = WaitForElementExists(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div/div[7]/div/div/div/nav/select"));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");

            return selectedValue == "30";
        }

        public bool isPageSizeEqualsTo50()
        {
            var nbPages = WaitForElementExists(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div/div[7]/div/div/div/nav/select"));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");

            return selectedValue == "50";
        }

        public bool isPageSizeEqualsTo100()
        {
            var nbPages = WaitForElementExists(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div/div[7]/div/div/div/nav/select"));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");

            return selectedValue == "100";
        }

        public int GetTotalRowsForPagination()
        {
            var table = _webDriver.FindElement(By.Id("tableDeliveries"));

            var rows = table.FindElements(By.TagName("tr"));

            // -1  Subtract 1 from the count to exclude the first row
            return rows.Count - 1;
        }

        public void SelectStatus(string statusLabel)
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions()
            {
                XpathId = STATUS_FILTER,
                SelectionValue = statusLabel,
                ClickCheckAllAtStart = false,
                IsUsedInFilter = false,
                ClickUncheckAllAtStart = false
            };
            ComboBoxSelectById(cbOpt);
        }
        public DeliveryMassiveDeleteRowResult GetRowResultInfo(int itemOffset)
        {
            string rowXpath = "(//*[@id=\"tableDeliveries\"]/tbody/tr)[" + (itemOffset + 1) + "]";
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
                DeliveryMassiveDeleteRowResult dsRowResult = new DeliveryMassiveDeleteRowResult();

                dsRowResult.DeliveryName = rowResult.FindElement(By.XPath(rowXpath + "/td[2]")).Text;

                if (string.IsNullOrEmpty(dsRowResult.DeliveryName) || dsRowResult.DeliveryName.StartsWith("No data"))
                { return null; }

                dsRowResult.IsDeliveryInactive = rowResult.FindElement(By.XPath(rowXpath + "/td[2]")).GetAttribute("class").Contains("IsInactive");
                dsRowResult.ServiceName = rowResult.FindElement(By.XPath(rowXpath + "/td[3]")).Text;
                dsRowResult.SiteName = rowResult.FindElement(By.XPath(rowXpath + "/td[4]")).Text;
                dsRowResult.IsSiteInactive = rowResult.FindElement(By.XPath(rowXpath + "/td[4]")).GetAttribute("class").Contains("IsInactive");
                dsRowResult.CustomerName = rowResult.FindElement(By.XPath(rowXpath + "/td[5]")).Text;
                dsRowResult.IsCustomerInactive = rowResult.FindElement(By.XPath(rowXpath + "/td[5]")).GetAttribute("class").Contains("IsInactive");
                dsRowResult.DeliveryRound = int.Parse(rowResult.FindElement(By.XPath(rowXpath + "/td[6]")).Text);

                string dateString = rowResult.FindElement(By.XPath(rowXpath + "/td[7]")).Text;
                if (string.IsNullOrEmpty(dateString) == false)
                {
                    dsRowResult.DateValidation = DateTime.ParseExact(dateString, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                }

                return dsRowResult;
            }
        }

        public void ShowInactiveSites()
        {
            var checkbox = WaitForElementExists(By.XPath(INACTIVE_SITE_CHECKBOX));
            checkbox.Click();
            WaitForLoad();
        }

        public void ClickOnSiteFilter()
        {
            var select = WaitForElementExists(By.Id(SITE_FILTER));
            select.Click();
            WaitForLoad();
        }

        public bool IsInactive()
        {
            var list = _webDriver.FindElements(By.XPath(SITE_SELECT_SEARCH)).ToList();

            var check = true;
            foreach (var element in list)
            {
                var text = element.Text;
                check = check && text.StartsWith("Inactive -");
            }
            WaitForLoad();
            return check;
        }
        public void ShowInactiveCustomers()
        {
            var checkbox = WaitForElementExists(By.XPath("//*/label[@for=\"ShowInactiveCustomers\"]/../div/input"));
            checkbox.Click();
            // cascade checkbox=>combobox
            WaitForLoad();
        }

        public void ClickOnCustomerFilter()
        {
            var select = WaitForElementExists(By.Id(CUSTOMER_FILTER));
            select.Click();
        }

        public bool CheckData()
        {
            try
            {
                var text = WaitForElementExists(By.XPath("//*[@id='tableDeliveries']/tbody/tr/td[2]/b")).Text;

                if (!text.Equals("No data to display. Use the Search button to load data")) return true;
                return false;
            }
            catch
            {
                return true;
            }

        }

        public int CheckTotalPageNumber()
        {
            try
            {
                return int.Parse(WaitForElementExists(By.XPath(MASSIVE_DELETE_PAGINATION)).Text);
            }
            catch
            {
                return 0;
            }
        }

        public void GoToPage(int page)
        {
            string paginationxpath = MASSIVE_DELETE_PAGINATION;
            if (!isElementExists(By.XPath(paginationxpath)))
            {
                paginationxpath = MASSIVE_DELETE_PAGINATION.Replace("/div[8]/div/div/nav/ul/li", "/div[8]/div/div/div/nav/ul/li");
            }
            var pageLink = _webDriver.FindElements(By.XPath(paginationxpath)).Where(e => e.Text.Trim().Equals(page.ToString())).FirstOrDefault();
            if (pageLink != null)
            {
                pageLink.Click();
                WaitForLoad();
            }
        }

        public bool VerifySiteSearch(string site)
        {
            var listSites = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_SERVICE_SITE_LIST));

            return listSites.Where(s => s.Text.Equals(site)).Count() == listSites.Count();
        }
        public void ClickDeleteButton()
        {
            _deleteDeliveriesBtn = WaitForElementIsVisible(By.Id(DELETEDELIVERIESBTN));
            _deleteDeliveriesBtn.Click();
            WaitPageLoading();
        }
        public void ClickConfirmDeleteButton()
        {
            _confirmationdelete = WaitForElementIsVisible(By.Id(CONFIRMATION_DELETE));
            _confirmationdelete.Click();
            WaitPageLoading();
        }
        public bool checkDeliveryExist()
        {
            return isElementExists(By.XPath(""));
        }
        public void UnSelectAll()
        {
            var unSelectAll = WaitForElementIsVisible(By.Id("unselectAll"));
            unSelectAll.Click();
        }
        public bool GetTotalTabs()
        {
            IList<string> allWindows = _webDriver.WindowHandles;

            return allWindows.Count > 1;
        }
        public int CheckTotalSelectCount()
        {
            var count = _webDriver.FindElement(By.Id(MASSIVE_DELETE_ITEM_SELECT_COUNT));
            var totalnumber = count.Text.Trim();
            return Convert.ToInt32(totalnumber);
        }
        public void SelectMultipleInactiveSites(string firstSite, string secondSite)
        {
            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", firstSite, false) );
            WaitPageLoading();
            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", secondSite, false) { ClickUncheckAllAtStart = false});
        }
        public void Cancel()
        {
            var unSelectAll = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[4]/button[1]"));
            unSelectAll.Click();
        }

        public class DeliveryMassiveDeleteRowResult
        {
            public string DeliveryName { get; set; }
            public bool IsDeliveryInactive { get; set; }
            public string ServiceName { get; set; }
            public string SiteName { get; set; }
            public bool IsSiteInactive { get; set; }
            public string CustomerName { get; set; }
            public bool IsCustomerInactive { get; set; }
            public int DeliveryRound { get; set; }
            public DateTime? DateValidation { get; set; }

        }
        public string GetMessageConfirmationMassive()
        {
            var _confirmationMessage = WaitForElementIsVisible(By.XPath(CONFIRMATIONMESSAGE));
            return _confirmationMessage.Text;
        }
    }
}