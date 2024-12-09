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
    public class EdiPurchaseOrdersPage : PageBase
    {
        private const string RESET_FILTER = "ResetFilter";

        private const string FILTER_SEARCH = "SearchPattern";
        private const string FILTER_SUPPLIER = "SelectedSupplierId";
        private const string FILTER_FROM = "date-picker-start";
        private const string FILTER_TO = "date-picker-end";
        private const string FILTER_SITES = "SelectedSites_ms";
        private const string FILTER_SITES_VALUE = "//*/button[@id='SelectedSites_ms']/span[2]";
        private const string FILTER_STATUS = "SelectedExportStatus_ms";
        private const string FILTER_STATUS_VALUE = "//*/button[@id='SelectedExportStatus_ms']/span[2]";
        private const string PURCHACE_ORDERS_DATES_LIST = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[5]";
        private const string PURCHACE_ORDER_FIRST_NUMBER_IN_LIST = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[2]/td[6]";
        private const string PURCHACE_ORDER_FIRST_SITENAME_LIST = "/html/body/div[2]/div/div[4]/div/div/div[2]/div[2]/table/tbody/tr[*]/td[3]/b";
        private const string EXPORT_BUTTON_ID = "btn-export-excel";

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
        [FindsBy(How = How.Id, Using = EXPORT_BUTTON_ID)]
        private IWebElement _exportPO;

        public enum FilterEdi
        {
            Search,
            Supplier,
            From,
            To,
            Sites,
            Status,
        }
        public EdiPurchaseOrdersPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
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
            }
            return null;
        }
        public IEnumerable<string> GetDatesListPurchaseOrders()
        {
            var listDatesPurchaseOrders = _webDriver.FindElements(By.XPath(PURCHACE_ORDERS_DATES_LIST));
            return listDatesPurchaseOrders.Select(p => p.Text);
        }
        public bool VerifyDatesPurchaseOrdersInFilter(IEnumerable<string> dates, DateTime from, DateTime to)
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

        public bool VerifyPurchaseOrdersToDate(IEnumerable<string> dates, DateTime to)
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

        public string GetfirstNumberPo()
        {
            var customerInvoiceNumber = _webDriver.FindElement(By.XPath(PURCHACE_ORDER_FIRST_NUMBER_IN_LIST));
            return customerInvoiceNumber.Text;
        }
        public IEnumerable<string> GetSiteNamesListPurchaseOrders()
        {
            var listDatesPurchaseOrders = _webDriver.FindElements(By.XPath(PURCHACE_ORDER_FIRST_SITENAME_LIST));
            return listDatesPurchaseOrders.Select(p => p.Text);
        }
        public void ExportExcelPO(bool versionPrint)
        {
            ShowExtendedMenu();
            _exportPO = WaitForElementIsVisible(By.Id(EXPORT_BUTTON_ID));
            _exportPO.Click();
            WaitForLoad();

            //no data, nothing to download
            //IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));

            //ClickPrintButton();
            //if (!isElementVisible(By.CssSelector("[target='_blank'][class='far fa-file-excel']")))
            //{
            //    Assert.IsTrue(true);//pas de données, pas de print, en attendant d'implémenter données
            //}

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            }
            //WaitForDownload();
            //Close();
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

            Regex reg = new Regex("^PO exports" + " " + annee + "-" + mois + "-" + jour + " " + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = reg.Match(filePath);

            return m.Success;
        }
    }
}
