using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Edi
{
    public class EdiPage : PageBase
    {
        private const string PURCHASE_ORDERS_TAB = "//*/li[@id='PurchaseOrderExport_ListItem']/a";
        private const string CUSTOMER_INVOICES_TAB = "//*/li[@id='CustomerInvoices_ListItem']/a";
        private const string CUSTOMER_ORDERS_TAB = "//*/li[@id='CustomerOrders_ListItem']/a";
        private const string SUPPLIER_INVOICES_TAB = "/html/body/div[2]/div/ul/li[1]/a";
        private const string RESET_FILTER = "ResetFilter";
        private const string SUPPLIER_INVOICES_DATES_LIST = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[6]";
        private const string FILTER_SEARCH = "SearchPattern";
        private const string FILTER_SUPPLIER = "SelectedSupplierId";
        private const string FILTER_FROM = "date-picker-start";
        private const string FILTER_TO = "date-picker-end";
        private const string FILTER_SITES = "SelectedSites_ms";
        private const string FILTER_SITES_VALUE = "//*/button[@id='SelectedSites_ms']/span[2]";
        private const string FILTER_STATUS = "SelectedImportStatus_ms";
        private const string FILTER_STATUS_VALUE = "//*/button[@id='SelectedImportStatus_ms']/span[2]";
        private const string FILTER_SHOW_ALL = "ShowOnlyValue";
        private const string FILTER_SHOW_ONLY_CN = "ShowOnlyValue";
        private const string IDS_COLUMN = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[2]/span";
        private const string DATES_COLUMN = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[6]";
        private const string EXTEND_MENU_EXPORT = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[1]/div/div[1]/button";
        private const string EXPORT_BUTTON = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[1]/div/div[1]/div/a[1]/span";
        private const string EXTEND_MENU_IMPORT = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[1]/div/div[2]/button";
        private const string IMPORT_BUTTON = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[1]/div/div[2]/div/a/span";
        private const string MODAL_IMPORT_OF_INVOICES = "/html/body/div[3]/div/div/div/div/form/div[1]";
        private const string INPUT_CHOOSE_FILE = "/html/body/div[3]/div/div/div/div/form/div[2]/div[1]/div/input";
        private const string CHECK_FILE_BUTTON = "/html/body/div[3]/div/div/div/div/form/div[3]/button[2]";
        private const string NUMBER_DATA = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[1]/h1/span";

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_SEARCH)]
        private IWebElement _searchFilter;
        [FindsBy(How = How.Id, Using = FILTER_SUPPLIER)]
        private IWebElement _supplierFilter;
        [FindsBy(How = How.Id, Using = FILTER_FROM)]
        private IWebElement _fromFilter;
        [FindsBy(How = How.Id, Using = FILTER_TO)]
        private IWebElement _toFilter;
        [FindsBy(How = How.Id, Using = FILTER_SITES)]
        private IWebElement _sitesFilter;
        [FindsBy(How = How.Id, Using = FILTER_STATUS)]
        private IWebElement _statusFilter;
        [FindsBy(How = How.Id, Using = FILTER_SHOW_ALL)]
        private IWebElement _showAllFilter;
        [FindsBy(How = How.Id, Using = FILTER_SHOW_ONLY_CN)]
        private IWebElement _showOnlyCNFilter;

        public enum FilterEdi
        {
            Search,
            Supplier,
            From,
            To,
            Sites,
            Status,
            ShowAll,
            ShowOnlyCN
        }

        public EdiPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public EdiPurchaseOrdersPage ClickOnPurchaseOrdersTab()
        {
            var tab = WaitForElementIsVisible(By.XPath(PURCHASE_ORDERS_TAB));
            tab.Click();
            WaitForLoad();
            return new EdiPurchaseOrdersPage(_webDriver, _testContext);
        }

        public EdiCustomerInvoicesPage ClickOnCustomerInvoicesTab()
        {
            var tab = WaitForElementIsVisible(By.XPath(CUSTOMER_INVOICES_TAB));
            tab.Click();
            WaitForLoad();
            return new EdiCustomerInvoicesPage(_webDriver, _testContext);
        }

        public EdiSupplierInvoicesPage ClickOnSupplierAccountingsTab()
        {
            var tab = WaitForElementIsVisible(By.XPath(SUPPLIER_INVOICES_TAB));
            tab.Click();
            WaitForLoad();
            return new EdiSupplierInvoicesPage(_webDriver, _testContext);
        }

        public EdiCustomerOrderPage ClickOnCustomerOrderTab()
        {
            var tab = WaitForElementIsVisible(By.XPath(CUSTOMER_ORDERS_TAB));
            tab.Click();
            WaitForLoad();
            return new EdiCustomerOrderPage(_webDriver, _testContext);
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

        public void Filter(FilterEdi FilterEdi, object value)
        {
            switch (FilterEdi)
            {
                case FilterEdi.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterEdi.Supplier:
                    _supplierFilter = WaitForElementIsVisible(By.Id(FILTER_SUPPLIER));
                    _supplierFilter.SetValue(ControlType.DropDownList, value);
                    break;
                case FilterEdi.From:
                    _fromFilter = WaitForElementIsVisible(By.Id(FILTER_FROM));
                    _fromFilter.SetValue(ControlType.TextBox, ((DateTime)value).ToString("dd/MM/yyyy"));
                    _fromFilter.SendKeys(Keys.Enter);
                    break;
                case FilterEdi.To:
                    _toFilter = WaitForElementIsVisible(By.Id(FILTER_TO));
                    _toFilter.SetValue(ControlType.TextBox, ((DateTime)value).ToString("dd/MM/yyyy"));
                    _toFilter.SendKeys(Keys.Enter);
                    break;
                case FilterEdi.Sites:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITES, (string)value));
                    break;
                case FilterEdi.Status:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_STATUS, (string)value));
                    break;
                case FilterEdi.ShowAll:
                    _showAllFilter = WaitForElementExists(By.Id(FILTER_SHOW_ALL));
                    _showAllFilter.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterEdi.ShowOnlyCN:
                    _showOnlyCNFilter = WaitForElementExists(By.Id(FILTER_SHOW_ONLY_CN));
                    _showOnlyCNFilter.SetValue(ControlType.RadioButton, value);
                    break;
            }
            WaitForLoad();
        }

        public object GetFilterValue(FilterEdi FilterEdi)
        {
            switch (FilterEdi)
            {
                case FilterEdi.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH));
                    return _searchFilter.GetAttribute("value");
                case FilterEdi.Supplier:
                    _supplierFilter = WaitForElementIsVisible(By.Id(FILTER_SUPPLIER));
                    var select = new SelectElement(_supplierFilter);
                    return select.SelectedOption.Text;
                case FilterEdi.From:
                    _fromFilter = WaitForElementIsVisible(By.Id(FILTER_FROM));
                    return _fromFilter.GetAttribute("value");
                case FilterEdi.To:
                    _toFilter = WaitForElementIsVisible(By.Id(FILTER_TO));
                    return _toFilter.GetAttribute("value");
                case FilterEdi.Sites:
                    _sitesFilter = WaitForElementIsVisible(By.XPath(FILTER_SITES_VALUE));
                    return _sitesFilter.Text;
                case FilterEdi.Status:
                    _statusFilter = WaitForElementIsVisible(By.XPath(FILTER_STATUS_VALUE));
                    return _statusFilter.Text;
                case FilterEdi.ShowAll:
                    _showAllFilter = WaitForElementIsVisible(By.Id(FILTER_SHOW_ALL));
                    return _showAllFilter.Selected;
                case FilterEdi.ShowOnlyCN:
                    _showOnlyCNFilter = WaitForElementIsVisible(By.Id(FILTER_SHOW_ONLY_CN));
                    return _showOnlyCNFilter.Selected;
            }
            return null;
        }

        public IEnumerable<string> GetListIds()
        {
            var ids = _webDriver.FindElements(By.XPath(IDS_COLUMN));
            return ids.Select(e => e.Text);
        }

        public void Export()
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU_EXPORT));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            WaitForLoad();
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

        public string[] GetNumberInString(string input)
        {
            string[] numbers = Regex.Matches(input, @"\d+").Cast<Match>().Select(m => m.Value).ToArray();
            return numbers;
        }

        public bool VerifyExcel(IEnumerable<string> idsGrid, List<string> idsExcel)
        {
            for (int i = 0; i < idsGrid.Count(); i++)
            {
                var number = GetNumberInString(idsGrid.ElementAt(i)).First();
                if (number != null)
                {
                    if (number != idsExcel[i])
                        return false;
                }
            }
            return true;
        }
        public void Import(string excelPath)
        {
            ClosePrintPopoverIfVisible();
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU_IMPORT));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            WaitForLoad();
            var importBtn = WaitForElementIsVisible(By.XPath(IMPORT_BUTTON));
            actions.MoveToElement(importBtn).Perform();
            importBtn.Click();
            
            var modalImportInvoices = WaitForElementIsVisible(By.XPath(MODAL_IMPORT_OF_INVOICES));
            modalImportInvoices.Click();
            var inputChooseFile = WaitForElementIsVisible(By.XPath(INPUT_CHOOSE_FILE));
            inputChooseFile.SendKeys(excelPath);
            var checkFileButton = WaitForElementIsVisible(By.XPath(CHECK_FILE_BUTTON));
            checkFileButton.Click();
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
        public IEnumerable<string> GetListStatuts()
        {
            var statutList = _webDriver.FindElements(By.XPath("/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[10]"));
            return statutList.Select(e => e.Text);
        }
        public bool VerifyStatutExcel(IEnumerable<string> coloneStatutGrid, List<string> statutExcel)
        {
            for (int i = 0; i < coloneStatutGrid.Count(); i++)
            {
                var statutColone = coloneStatutGrid.ElementAt(i);
                if (statutColone != null)
                {
                    if (statutColone != statutExcel[i])
                        return false;
                }
            }
            return true;
        }
    }
}
