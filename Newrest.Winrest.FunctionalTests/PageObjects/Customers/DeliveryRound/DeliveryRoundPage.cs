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
using System.Text.RegularExpressions;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.DeliveryRound
{
    public class DeliveryRoundPage : PageBase
    {
        public DeliveryRoundPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_______________________________________

        // General
        private const string EXTENDED_BUTTON = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/button";
        private const string EXPORT = "btn-export-excel";
        private const string UNFOLD_ALL = "//*[@id=\"unfoldBtn\"][@class='btn unfold-all-btn']";

        private const string CREATE_NEW_DELIVERY_ROUND = "New Delivery round";

        // Tableau
        private const string FIRST_DELIVERY_ROUND = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[2]";

        // Filtres
        private const string RESET_FILTER = "//a[contains(text(), 'Reset Filter')]";

        private const string SEARCH_FILTER = "SearchPattern";
        private const string SORTBY_FILTER = "cbSortBy";

        private const string SHOW_ALL_FILTER = "ShowAll";
        private const string SHOW_ONLY_ACTIVE_FILTER = "ShowActive";
        private const string SHOW_ONLY_INACTIVE_FILTER = "ShowInactive";

        private const string DATE_FROM = "date-picker-start";
        private const string DELETE_FIRST = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[3]/a";
        private const string DELETE_CONFIRM = "dataConfirmOK";
        private const string DATE_TO = "date-picker-end";
        private const string SITE_FILTER = "SelectedSites_ms";
        private const string SITES_UNSELECT_ALL = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SEARCH_SITE = "/html/body/div[10]/div/div/label/input";
        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string CUSTOMER_UNSELECT_ALL = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_CUSTOMER = "/html/body/div[11]/div/div/label/input";

        private const string INACTIVE = "//*[@id=\"list-item-with-action\"]/div[{0}]/div/div/div[2]/table/tbody/tr/td[1]/img[@alt=\"Inactive\"]";
        private const string DELIVERY_NAMES = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string PERIOD = "//*[@id=\"list-item-with-action\"]/div[*]/div/div/div[2]/table/tbody/tr/td[3]";
        private const string SITES = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[4]";
        private const string CUSTOMERS = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[2]";
        private const string PLUS_BUTTON = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[2]/button";
        private const string DELETE_BTN = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[2]/div[1]/div/div[3]/a/span";
        private const string CONFIRM_DELETE_BTN = "dataConfirmOK";
        private const string ROW_TO_HOVER = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]";

        //__________________________________Variables_______________________________________

        // General
        [FindsBy(How = How.XPath, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.XPath, Using = UNFOLD_ALL)]
        private IWebElement _unfoldAll;

        [FindsBy(How = How.LinkText, Using = CREATE_NEW_DELIVERY_ROUND)]
        private IWebElement _createNewDeliveryRound;

        // Tableau
        [FindsBy(How = How.XPath, Using = FIRST_DELIVERY_ROUND)]
        private IWebElement _firstDeliveryRound;

        // __________________________________________________ Filtres __________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = SHOW_ALL_FILTER)]
        private IWebElement _showAll;

        [FindsBy(How = How.Id, Using = SHOW_ONLY_ACTIVE_FILTER)]
        private IWebElement _showOnlyActiveDev;

        [FindsBy(How = How.Id, Using = SHOW_ONLY_INACTIVE_FILTER)]
        private IWebElement _showOnlyInactive;

        [FindsBy(How = How.Id, Using = DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = SITES_UNSELECT_ALL)]
        private IWebElement _siteUnselectAll;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.Id, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = CUSTOMER_UNSELECT_ALL)]
        private IWebElement _customerUnselectAll;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.XPath, Using = PLUS_BUTTON)]
        private IWebElement _plusButton;

        public enum FilterType
        {
            Search,
            SortBy,
            ShowAll,
            ShowOnlyActive,
            ShowOnlyInactive,
            DateFrom,
            DateTo,
            Site,
            Customers
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORTBY_FILTER));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;

                case FilterType.ShowAll:
                    _showAll = WaitForElementExists(By.Id(SHOW_ALL_FILTER));
                    _showAll.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowOnlyActive:_showOnlyActiveDev = WaitForElementExists(By.Id(SHOW_ONLY_ACTIVE_FILTER));
                    _showOnlyActiveDev.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowOnlyInactive:
                    _showOnlyInactive = WaitForElementExists(By.Id(SHOW_ONLY_INACTIVE_FILTER));
                    _showOnlyInactive.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.DateFrom:
                    _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;
                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;

                case FilterType.Site:
                    ComboBoxSelectById(new ComboBoxOptions(SITE_FILTER, (string)value));
                    break;
                case FilterType.Customers:
                    _customerFilter = WaitForElementIsVisible(By.Id(CUSTOMER_FILTER));
                    _customerFilter.Click();

                    _customerUnselectAll = WaitForElementIsVisible(By.XPath(CUSTOMER_UNSELECT_ALL));
                    _customerUnselectAll.Click();

                    _searchCustomer = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER));
                    _searchCustomer.SetValue(ControlType.TextBox, value);
                    Thread.Sleep(1500);

                    var valueToCheckCustomers = WaitForElementIsVisible(By.XPath("/html/body/div[11]/ul"));
                    valueToCheckCustomers.Click();

                    _customerFilter.Click();
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

        public bool IsSortedByName()
        {
            var elements = _webDriver.FindElements(By.XPath(DELIVERY_NAMES));

            if (elements.Count == 0)
                return false;

            var names = elements.Select(elm => elm.Text).ToList();

            for (int i = 1; i < names.Count; i++)
            {
                if (string.Compare(names[i], names[i - 1]) < 0)
                    return false;
            }

            return true;
        }

        public bool IsSortedBySite()
        {
            var listeSites = _webDriver.FindElements(By.XPath(SITES));

            if (listeSites.Count == 0)
                return false;

            var ancienSite = listeSites[0].Text;

            foreach (var elm in listeSites)
            {
                try
                {
                    if (elm.Text.CompareTo(ancienSite) < 0)
                        return false;

                    ancienSite = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool CheckStatus(bool active)
        {
            bool valueBool = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 1; i <= tot; i++)
            {
                if (isElementVisible(By.XPath(String.Format(INACTIVE, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(String.Format(INACTIVE, i + 1)));

                    if (active)
                        return false;
                }
                else
                {
                    valueBool = true;
                    if (!active)
                        return true;
                }
            }
            return valueBool;
        }

        public bool IsDateRespected(DateTime fromDate, DateTime toDate, string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var dates = _webDriver.FindElements(By.XPath(PERIOD));

            if (dates.Count == 0)
                return false;

            foreach (var elm in dates)
            {
                try
                {
                    int separator = elm.Text.IndexOf("-");

                    string startDateText = elm.Text.Substring(0, separator).Trim();
                    string endDateText = elm.Text.Substring(separator + 1).Trim();

                    DateTime startDate = DateTime.Parse(startDateText, ci).Date;
                    DateTime endDate = DateTime.Parse(endDateText, ci).Date;

                    if (DateTime.Compare(startDate, fromDate) < 0 || DateTime.Compare(endDate, fromDate) < 0 || DateTime.Compare(endDate, toDate) > 0 || DateTime.Compare(startDate, toDate) > 0)
                        return false;
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifySite(string site)
        {
            var sites = _webDriver.FindElements(By.XPath(SITES));

            if (sites.Count == 0)
                return false;

            foreach (var elm in sites)
            {
                if (!elm.Text.Contains(site))
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifyCustomer(string value)
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 1; i <= tot; i++)
            {
                try
                {
                    bool isCustomerFound = false;
                    var elements = _webDriver.FindElements(By.XPath(string.Format(CUSTOMERS, i + 1)));

                    foreach (var elm in elements)
                    {
                        if (elm.Text.Contains(value))
                        {
                            isCustomerFound = true;
                            break;
                        }
                    }

                    if (!isCustomerFound)
                        return false;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool VerifyDRNamesExist(string drName)
        {
            List<string> strings = new List<string>();
            var names = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[2]"));
            foreach (var name in names)
            {
                strings.Add(name.Text);
            }

            if(!strings.Contains(drName))
            {
                return false;
            }
            return true;
        }

        public void DeleteFirstDeliveryRound()
        {
            var actions = new Actions(_webDriver);
            var deleteFirst = WaitForElementExists(By.XPath(DELETE_FIRST));
            actions.MoveToElement(deleteFirst).Perform();
            deleteFirst.Click();

            var deleteConfirm = WaitForElementExists(By.Id(DELETE_CONFIRM));
            deleteConfirm.Click();
            WaitForLoad();
        }

        // __________________________________________________ Méthodes _________________________________________________

        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);

            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BUTTON));
            actions.MoveToElement(_extendedButton).Perform();
        }

        public void Export(bool versionPrint)
        {
            ShowExtendedMenu();
            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles, string deliveryRoundName = null)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                if (IsExcelFileCorrect(file.Name, deliveryRoundName))
                {
                    correctDownloadFiles.Add(file);
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

        public bool IsExcelFileCorrect(string filePath, string deliveryRoundName)
        {
            Regex reg;

            if (deliveryRoundName != null)
            {
                reg = new Regex("^export-rounddelivery" + "-" + deliveryRoundName.Replace("/", "_") + ".xlsx", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            }
            else
            {
                reg = new Regex("^export-rounddelivery(.*).xlsx", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            }

            Match m = reg.Match(filePath);

            return m.Success;
        }

        public void UnfoldAll()
        {
            ShowExtendedMenu();

            _unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
            _unfoldAll.Click();
            WaitForLoad();
        }


        public DeliveryRoundCreateModalpage DeliveryRoundCreatePage()
        {
            ShowPlusMenu();

            _createNewDeliveryRound = WaitForElementIsVisible(By.LinkText(CREATE_NEW_DELIVERY_ROUND));
            _createNewDeliveryRound.Click();
            WaitForLoad();

            return new DeliveryRoundCreateModalpage(_webDriver, _testContext);
        }

        // Tableau
        public string GetFirstDeliveryRound()
        {
            _firstDeliveryRound = WaitForElementIsVisible(By.XPath(FIRST_DELIVERY_ROUND));
            return _firstDeliveryRound.Text;
        }

        public DeliveryRoundDeliveryPage SelectFirstDeliveryRound()
        {
            _firstDeliveryRound = WaitForElementIsVisible(By.XPath(FIRST_DELIVERY_ROUND));
            _firstDeliveryRound.Click();
            WaitForLoad();

            return new DeliveryRoundDeliveryPage(_webDriver, _testContext);
        }

        public bool AccessPage()
        {
            if (isElementExists(By.XPath(PLUS_BUTTON)))
            {
                return true;
            }

            return false;
        }
        public void DeleteDelivery()
        {
            
                // Locate the row or element to hover over
                var _rowToHover = _webDriver.FindElement(By.XPath(ROW_TO_HOVER));

                var actions = new OpenQA.Selenium.Interactions.Actions(_webDriver);
                actions.MoveToElement(_rowToHover).Perform();

                var _deleteBtn = _webDriver.FindElement(By.XPath(DELETE_BTN));

                _deleteBtn.Click();
                WaitForLoad();

                var _confirmDeleteBtn = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_BTN));

                _confirmDeleteBtn.Click();
                WaitForLoad();
            
       
        }
        public void DeleteDeliveryRound()
        {

            // Locate the row or element to hover over
            var _rowToHover = _webDriver.FindElement(By.XPath(ROW_TO_HOVER));

            var actions = new OpenQA.Selenium.Interactions.Actions(_webDriver);
            actions.MoveToElement(_rowToHover).Perform();

            var _deleteBtn = _webDriver.FindElement(By.XPath(DELETE_BTN));

            _deleteBtn.Click();
            WaitForLoad();

            var _confirmDeleteBtn = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_BTN));

            _confirmDeleteBtn.Click();
            WaitForLoad();
        }

    }
}
