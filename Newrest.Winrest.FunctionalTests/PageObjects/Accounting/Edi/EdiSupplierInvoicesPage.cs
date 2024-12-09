using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Edi
{
    public class EdiSupplierInvoicesPage : PageBase
    {
        public EdiSupplierInvoicesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
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
        private const string SHOW_ONLY_CN = "input[value='CreditNotesOnly']";
        private const string SHOW_ONLY_SI ="input[value='SupplierInvoicesOnly']";
        private const string IDS_COLUMN = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[2]/span";
        private const string DATES_COLUMN = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[6]";
        private const string EXTEND_MENU_EXPORT = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[1]/h1/div/div[1]/button";
        private const string EXPORT_BUTTON = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[1]/h1/div/div[1]/div/a[1]";
        private const string EXTEND_MENU_IMPORT = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[1]/h1/div/div[2]/button";
        private const string IMPORT_BUTTON = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[1]/h1/div/div[2]/div/a";
        private const string MODAL_IMPORT_OF_INVOICES = "/html/body/div[3]/div/div/div/div/form/div[1]";
        private const string INPUT_CHOOSE_FILE = "/html/body/div[3]/div/div/div/div/form/div[2]/div[1]/div/input";
        private const string CHECK_FILE_BUTTON = "/html/body/div[3]/div/div/div/div/form/div[3]/button[2]";
        private const string NUMBER_DATA = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[1]/h1/span";
        private const string STATUS = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[10]";

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

        [FindsBy(How = How.XPath, Using = STATUS)]
        public IWebElement _status;

        public enum FilterType
        {
            Search,
            Supplier,
            From,
            To,
            Sites,
            Status,
            ShowAll,
            ShowOnlyCN,
            ShowOnlySI
        }


        public void Filter(FilterType FilterType, object value)
        {
            switch (FilterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    WaitPageLoading();
                    break;
                case FilterType.Supplier:
                    _supplierFilter = WaitForElementIsVisible(By.Id(FILTER_SUPPLIER));
                    _supplierFilter.SetValue(ControlType.DropDownList, value);
                    WaitPageLoading();
                    break;
                case FilterType.From:
                    _fromFilter = WaitForElementIsVisible(By.Id(FILTER_FROM));
                    _fromFilter.SetValue(ControlType.TextBox, ((DateTime)value).ToString("dd/MM/yyyy"));
                    _fromFilter.SendKeys(Keys.Tab);
                    WaitPageLoading();
                    break;
                case FilterType.To:
                    _toFilter = WaitForElementIsVisible(By.Id(FILTER_TO));
                    _toFilter.SetValue(ControlType.TextBox, ((DateTime)value).ToString("dd/MM/yyyy"));
                    _toFilter.SendKeys(Keys.Tab);
                    WaitPageLoading();
                    break;
                case FilterType.Sites:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITES, (string)value));
                    WaitPageLoading();
                    break;
                case FilterType.Status:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_STATUS, (string)value));
                    WaitPageLoading();
                    break;
                case FilterType.ShowAll:
                    _showAllFilter = WaitForElementExists(By.Id(FILTER_SHOW_ALL));
                    _showAllFilter.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowOnlyCN:
                    _showOnlyCNFilter = WaitForElementExists(By.CssSelector(SHOW_ONLY_CN));
                    _showOnlyCNFilter.SetValue(ControlType.RadioButton, value);            
                    WaitPageLoading();
                    break;
                case FilterType.ShowOnlySI:
                    _showOnlyCNFilter = WaitForElementExists(By.CssSelector(SHOW_ONLY_SI));
                    _showOnlyCNFilter.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;
            }
            WaitForLoad();
        }

        public object GetFilterValue(FilterType FilterType)
        {
            switch (FilterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH));
                    return _searchFilter.GetAttribute("value");
                case FilterType.Supplier:
                    _supplierFilter = WaitForElementIsVisible(By.Id(FILTER_SUPPLIER));
                    var select = new SelectElement(_supplierFilter);
                    return select.SelectedOption.Text;
                case FilterType.From:
                    _fromFilter = WaitForElementIsVisible(By.Id(FILTER_FROM));
                    return _fromFilter.GetAttribute("value");
                case FilterType.To:
                    _toFilter = WaitForElementIsVisible(By.Id(FILTER_TO));
                    return _toFilter.GetAttribute("value");
                case FilterType.Sites:
                    _sitesFilter = WaitForElementIsVisible(By.XPath(FILTER_SITES_VALUE));
                    return _sitesFilter.Text;
                case FilterType.Status:
                    _statusFilter = WaitForElementIsVisible(By.XPath(FILTER_STATUS_VALUE));
                    return _statusFilter.Text;
                case FilterType.ShowAll:
                    _showAllFilter = WaitForElementIsVisible(By.Id(FILTER_SHOW_ALL));
                    return _showAllFilter.Selected;
                case FilterType.ShowOnlyCN:
                    _showOnlyCNFilter = WaitForElementIsVisible(By.Id(FILTER_SHOW_ONLY_CN));
                    return _showOnlyCNFilter.Selected;
            }
            return null;
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
        public int GetNumberSupplierInvoicesEdi()
        {
            var number = WaitForElementIsVisible(By.XPath(NUMBER_DATA));
            return int.Parse(number.Text.ToString());
        }
        public IEnumerable<string> GetDatesListSupplierInvoices()
        {
            var listDatesSupplierInvocies = _webDriver.FindElements(By.XPath(SUPPLIER_INVOICES_DATES_LIST));
            return listDatesSupplierInvocies.Select(p => p.Text);
        }
        public bool VerifyDatesSupplierInvoicesInFilter(IEnumerable<string> dates, DateTime from, DateTime to)
        {
            foreach (var date in dates)
            {
                if (DateTime.ParseExact(date, "dd/MM/yyyy", new CultureInfo("fr-FR")) > to || (DateTime.ParseExact(date, "dd/MM/yyyy", new CultureInfo("fr-FR")) < from))
                {
                    return false;
                }
            }
            return true;
        }
        public bool VerifySupplierInvoicesToDate(IEnumerable<string> dates, DateTime to)
        {

            foreach (var date in dates)
            {
                if (DateTime.ParseExact(date, "dd/MM/yyyy", new CultureInfo("fr-FR")) > to)
                {
                    return false;
                }
            }
            return true;
        }
        public bool IsAddedFileVerif(int numberRawsListSupplierInvoices)
        {
            var numberListSupplierInvoices = GetNumberSupplierInvoicesEdi();
            if (numberListSupplierInvoices != numberRawsListSupplierInvoices + 1)
            {
                return false;
            }
            var Status = WaitForElementIsVisible(By.XPath("//tr[2]//span[@class='text-danger']")).Text;
            if (Status.Contains(".xlsx not supported"))
            {
                return false;
            }

            return true;
        }


        public bool IsListeStatusAffiche()
        { 
            var status = _webDriver.FindElements(By.XPath(STATUS));

            if (status.Count == 0)
                return false;

            return status.Any(c => !string.IsNullOrEmpty(c.Text));
        }
 
        public EdiSupplierInvoicesPage ClickOnFirstItem()
        {
            WaitLoading();
            if(isElementExists(By.XPath("/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[2]/td[1]")))
            {

                var _clickOnElement = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[2]/td[1]"));
                _clickOnElement.Click();
            }
            return new EdiSupplierInvoicesPage(_webDriver, _testContext);
        }

        public bool IsRetourligneFormatEDI(string format, string fileName, string fileParsing)
        {
            var formatt = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[3]/p[1]")).Text;
            var fileNamee = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[3]/p[2]")).Text;
            var fileParsingg = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[4]/div/div/p")).Text;

            if ((formatt.Contains(format)) && (fileNamee.Contains(fileName)) && (fileParsingg.Contains(fileParsing)))
            {
                return true;
            }

           else  return false;

      
        }
    }
}
