using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Edi
{
    public class EdiCustomerInvoicesPage : PageBase
    {

        private const string RESET_FILTER = "ResetFilter";

        private const string FILTER_SEARCH = "SearchPattern";
        private const string FILTER_CUSTOMER = "SelectedCustomerId";
        private const string FILTER_FROM = "date-picker-start";
        private const string FILTER_TO = "date-picker-end";
        private const string FILTER_SITES = "SelectedSites_ms";
        private const string FILTER_SITES_VALUE = "//*/button[@id='SelectedSites_ms']/span[2]";
        private const string FILTER_STATUS = "SelectedExportStatus_ms";
        private const string FILTER_STATUS_VALUE = "//*/button[@id='SelectedExportStatus_ms']/span[2]";
        private const string CUSTOMER_INVOICES_DATES_LIST = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[7]";
        private const string FIRST_CUSTOMER_INVOICE_INVOICE_NUMBER = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[2]/td[6]";
        private const string FIRST_CUSTOMER_INVOICE_SITE_NAME = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[2]/td[3]/b";
        private const string CUSTOMER_INVOICE_SITE_NAME_LIST = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[3]/b";

        private const string FIRST_CUSTOMER_INVOICE_STATUS = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[2]/td[8]/span";
        private const string CUSTOMER_INVOICE_STATUS_LIST = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[8]/span";
        private const string EXPORT_BUTTON_ID = "btn-export-excel";

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_SEARCH)]
        private IWebElement _searchFilter;
        [FindsBy(How = How.Id, Using = FILTER_CUSTOMER)]
        private IWebElement _customerFilter;
        [FindsBy(How = How.Id, Using = FILTER_FROM)]
        private IWebElement _fromFilter;
        [FindsBy(How = How.Id, Using = FILTER_TO)]
        private IWebElement _toFilter;
        [FindsBy(How = How.Id, Using = FILTER_SITES)]
        private IWebElement _sitesFilter;
        [FindsBy(How = How.Id, Using = FILTER_STATUS)]
        private IWebElement _statusFilter;
        [FindsBy(How = How.Id, Using = EXPORT_BUTTON_ID)]
        private IWebElement _exportCI;

        public enum FilterEdi
        {
            Search,
            Customer,
            From,
            To,
            Sites,
            Status,
        }

        public EdiCustomerInvoicesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
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
                case FilterEdi.Customer:
                    _customerFilter = WaitForElementIsVisible(By.Id(FILTER_CUSTOMER));
                    _customerFilter.SetValue(ControlType.DropDownList, value);
                    break;
                case FilterEdi.From:
                    _fromFilter = WaitForElementIsVisible(By.Id(FILTER_FROM));
                    _fromFilter.SetValue(ControlType.TextBox, ((DateTime)value).ToString("dd/MM/yyyy"));
                    _fromFilter.SendKeys(Keys.Enter);
                    _fromFilter.SendKeys(Keys.Tab);
                    break;
                case FilterEdi.To:
                    _toFilter = WaitForElementIsVisible(By.Id(FILTER_TO));
                    _toFilter.SetValue(ControlType.TextBox, ((DateTime)value).ToString("dd/MM/yyyy"));
                    _toFilter.SendKeys(Keys.Enter);
                    _toFilter.SendKeys(Keys.Tab);
                    break;
                case FilterEdi.Sites:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITES, (string)value));
                    break;
                case FilterEdi.Status:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_STATUS, (string)value));
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
                case FilterEdi.Customer:
                    _customerFilter = WaitForElementIsVisible(By.Id(FILTER_CUSTOMER));
                    var select = new SelectElement(_customerFilter);
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
            }
            return null;
        }

        public IEnumerable<string> GetDatesListCustomerInvoices()
        {
            var listDatesCustomerInvoices = _webDriver.FindElements(By.XPath(CUSTOMER_INVOICES_DATES_LIST));
            return listDatesCustomerInvoices.Select(p => p.Text);
        }
        public bool VerifyDatesCustomerInvoicesInFilter(IEnumerable<string> dates, DateTime from, DateTime to)
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
        public bool VerifyDatesCustomerInvoicesFilterFrom(IEnumerable<string> dates, DateTime from)
        {
            foreach (var date in dates)
            {
                if (DateTime.ParseExact(date, "dd/MM/yyyy", new CultureInfo("fr-FR")) < from)
                {
                    return false;
                }
            }
            return true;
        }
        public bool VerifyDatesCustomerInvoicesFilterTo(IEnumerable<string> dates, DateTime to)
        {
            foreach (var date in dates)
            {
                if (DateTime.ParseExact(date, "dd/MM/yyyy", new CultureInfo("fr-FR")) > to )
                {
                    return false;
                }
            }
            return true;
        }
        public string GetFirstCIInvoiceNumber()
        {
            var invoiceNumber = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER_INVOICE_INVOICE_NUMBER));
            return invoiceNumber.Text;
        }
        public string GetFirstCISiteName()
        {
            var invoiceNumber = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER_INVOICE_SITE_NAME));
            return invoiceNumber.Text;
        }
        public IEnumerable<string> GetListSiteName()
        {
            var listDatesCustomerInvoices = _webDriver.FindElements(By.XPath(CUSTOMER_INVOICE_SITE_NAME_LIST));
            return listDatesCustomerInvoices.Select(p => p.Text);
        }

        public string GetFirstCIStatus()
        {
            var invoiceNumber = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER_INVOICE_STATUS));
            return invoiceNumber.Text;
        }
        public IEnumerable<string> GetListStatus()
        {
            var listDatesCustomerInvoices = _webDriver.FindElements(By.XPath(CUSTOMER_INVOICE_STATUS_LIST));
            return listDatesCustomerInvoices.Select(p => p.Text);
        }
        public void ExportExcelCI(bool versionPrint)
        {

            ShowExtendedMenu();
            _exportCI = WaitForElementIsVisible(By.Id(EXPORT_BUTTON_ID));
            _exportCI.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            }

            WaitForDownload();
            Close();
        }
        public FileInfo GetExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
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
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // minutes

            Regex reg = new Regex("^CI exports" + " " + annee + "-" + mois + "-" + jour + " " + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = reg.Match(filePath);

            return m.Success;
        }
    }
}
