using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reinvoice

{


    public class ReinvoicePage : PageBase
    {

        public ReinvoicePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________ Cosntantes ______________________________________________________

        // General
        private const string NEW_REINVOICE = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[2]/div/a[1]";
        private const string NEW_REINVOICE_WITH_SITES = "New reinvoice with sites";
        private const string EXPORT = "btn-export-excel";

        // Tableau
        private const string CUSTOMER_FROM = "//*[@id=\"reinvoice-table\"]/tbody/tr[*]/td[2]";
        private const string CUSTOMER_FROM_PATCH = "//*[@id=\"reinvoice-table\"]/tbody/tr[*]/td[1]";
        private const string CUSTOMER_TO = "//*[@id=\"reinvoice-table\"]/tbody/tr[*]/td[5]";
        private const string CUSTOMER_TO_PATCH = "//*[@id=\"reinvoice-table\"]/tbody/tr[*]/td[4]";
        private const string DELETE = "//span[contains(@class, 'glyphicon glyphicon-trash')]";
        private const string CONFIRM_DELETE = "dataConfirmOK";

        // Filtres
        private const string RESET_FILTER = "ResetFilter";
        private const string CUSTOMER_FROM_FILTER = "SelectedCustomersFrom_ms";
        private const string UNCHECK_ALL_CUSTOMER_FROM = "/html/body/div[10]/div/ul/li[2]/a";
        private const string CHECK_ALL_CUSTOMER_FROM = "/html/body/div[10]/div/ul/li[1]/a";
        private const string CUSTOMER_FROM_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string CUSTOMER_TO_FILTER = "SelectedCustomersTo_ms";
        private const string UNCHECK_ALL_CUSTOMER_TO = "/html/body/div[11]/div/ul/li[2]/a";
        private const string CHECK_ALL_CUSTOMER_TO = "/html/body/div[11]/div/ul/li[1]/a";
        private const string CUSTOMER_TO_SEARCH = "/html/body/div[11]/div/div/label/input";

        private const string NB_CUSTOMER_FROM_SELECTED = "//*[@id=\"SelectedCustomersFrom_ms\"]/span[2]";
        private const string NB_CUSTOMER_TO_SELECTED = "//*[@id=\"SelectedCustomersTo_ms\"]/span[2]";


        // _____________________________________ Variables _______________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_REINVOICE)]
        private IWebElement _createNewReinvoice;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        // Tableau
        [FindsBy(How = How.XPath, Using = CUSTOMER_FROM)]
        private IWebElement _firstCustomerFrom;
        
        [FindsBy(How = How.XPath, Using = DELETE)]
        private IWebElement _delete;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        // _____________________________________ Filtres _______________________________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.XPath, Using = CUSTOMER_FROM_FILTER)]
        private IWebElement _customersFromFilter;

        [FindsBy(How = How.XPath, Using = CUSTOMER_TO)]
        private IWebElement _firstCustomerTo;
        

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_CUSTOMER_FROM)]
        private IWebElement _unselectAllCustomerFrom;

        [FindsBy(How = How.XPath, Using = CHECK_ALL_CUSTOMER_FROM)]
        private IWebElement _selectAllCustomerFrom;

        [FindsBy(How = How.XPath, Using = CUSTOMER_FROM_SEARCH)]
        private IWebElement _searchCustomerFrom;

        [FindsBy(How = How.Id, Using = CUSTOMER_TO_FILTER)]
        private IWebElement _customersToFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_CUSTOMER_TO)]
        private IWebElement _unselectAllCustomerTo;

        [FindsBy(How = How.XPath, Using = CHECK_ALL_CUSTOMER_TO)]
        private IWebElement _selectAllCustomerTo;

        [FindsBy(How = How.XPath, Using = CUSTOMER_TO_SEARCH)]
        private IWebElement _searchCustomerTo;

        public enum FilterType
        {
            SearchCustomersFrom,
            SearchCustomersTo
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.SearchCustomersFrom:
                    _customersFromFilter = WaitForElementIsVisible(By.Id(CUSTOMER_FROM_FILTER));
                    _customersFromFilter.Click();

                    WaitForLoad();

                    _unselectAllCustomerFrom = _webDriver.FindElement(By.XPath(UNCHECK_ALL_CUSTOMER_FROM));
                    _unselectAllCustomerFrom.Click();
                    WaitPageLoading();

                    _searchCustomerFrom = WaitForElementIsVisible(By.XPath(CUSTOMER_FROM_SEARCH));
                    _searchCustomerFrom.SetValue(ControlType.TextBox, value);

                    Thread.Sleep(2000);//waitforload pas suffisant

                    if(isElementVisible(By.XPath("//span[text()='" + value + "']")))
                    {
                        var customerFromToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                        customerFromToCheck.Click();
                        WaitPageLoading();
                    }
                    else
                    {
                        _selectAllCustomerFrom = _webDriver.FindElement(By.XPath(CHECK_ALL_CUSTOMER_FROM));
                        _selectAllCustomerFrom.Click();
                    }

                    _customersFromFilter.Click();
                    WaitForLoad();

                    break;

                case FilterType.SearchCustomersTo:
                    _customersToFilter = WaitForElementIsVisible(By.Id(CUSTOMER_TO_FILTER));
                    _customersToFilter.Click();
                    WaitForLoad();

                    _unselectAllCustomerTo = _webDriver.FindElement(By.XPath(UNCHECK_ALL_CUSTOMER_TO));
                    _unselectAllCustomerTo.Click();
                    WaitPageLoading();

                    _searchCustomerTo = WaitForElementIsVisible(By.XPath(CUSTOMER_TO_SEARCH));
                    _searchCustomerTo.SetValue(ControlType.TextBox, value);

                    Thread.Sleep(2000);//waitforload pas suffisant

                    if(isElementVisible(By.XPath("/html/body/div[11]/ul")))
                    {
                        var customerToCheck = _webDriver.FindElement(By.XPath("/html/body/div[11]/ul"));
                        customerToCheck.Click();
                        WaitPageLoading();
                    }
                    else
                    {
                        _selectAllCustomerTo = _webDriver.FindElement(By.XPath(CHECK_ALL_CUSTOMER_TO));
                        _selectAllCustomerTo.Click();
                        WaitForLoad();
                    }

                    _customersToFilter.Click();
                    WaitForLoad();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();
            //WaitForLoad();//waitforload sur cette page plante d'où le thread.sleep
            Thread.Sleep(8000);

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public bool VerifyCustomerFrom(string customerFrom)
        {
            var customersFrom = _webDriver.FindElements(By.XPath(CUSTOMER_FROM));

            if (customersFrom.Count == 0)
                return false;

            foreach (var elm in customersFrom)
            {
                if (!elm.Text.Contains(customerFrom))
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifyCustomerTo(string customerTo)
        {
            var customersTo = _webDriver.FindElements(By.XPath(CUSTOMER_TO));

            if (customersTo.Count == 0)
                return false;

            foreach (var elm in customersTo)
            {
                if (!elm.Text.Contains(customerTo))
                {
                    return false;
                }
            }
            return true;
        }

        // ______________________________________ Méthodes ___________________________________________________________       

        // General
        public ReinvoiceCreateModalPage CreateNewReinvoice()
        {
            ShowPlusMenu();

            _createNewReinvoice = WaitForElementIsVisible(By.XPath(NEW_REINVOICE));
            _createNewReinvoice.Click();
            WaitForLoad();

            return new ReinvoiceCreateModalPage(_webDriver, _testContext);
        }
        public ReinvoiceCreateModalPage CreateNewReinvoiceWithSites()
        {
            ShowPlusMenu();

            _createNewReinvoice = WaitForElementIsVisible(By.LinkText(NEW_REINVOICE_WITH_SITES));
            _createNewReinvoice.Click();
            WaitForLoad();

            return new ReinvoiceCreateModalPage(_webDriver, _testContext);
        }

        public void Export(bool versionPrint)
        {
            if (isElementVisible(By.XPath("//h3[contains(text(), 'Print list')]")))
            {
                ClickPrintButton();
            }
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
            //Close();
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                // Test REGEX
                if (IsExcelFileCorrect(file.Name))
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

        public bool IsExcelFileCorrect(string filePath)
        {
            // export-reinvoice-2020-05-29 12-29-54.xlsx

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            Regex r = new Regex("^export-reinvoice-" + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        // Tableau
        public void DeleteReinvoice()
        {
            WaitForLoad();
            _delete = WaitForElementIsVisible(By.XPath("//span[contains(@class, 'trash')]"));
            _delete.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        public string GetNumberCustFromSelected()
        {
            var element = _webDriver.FindElement(By.XPath(NB_CUSTOMER_FROM_SELECTED));
            return element.GetAttribute("innerText");
        }

        public string GetNumberCustToSelected()
        {
            var element = _webDriver.FindElement(By.XPath(NB_CUSTOMER_TO_SELECTED));
            return element.GetAttribute("innerText");
        }

        public string GetFirstCustomerFromName()
        {
            WaitForLoad();
            int code;
            if (CheckTotalNumber() == 0)
            {
                return null;
            }
            _firstCustomerFrom = WaitForElementIsVisible(By.XPath(CUSTOMER_FROM));
            code = _firstCustomerFrom.Text.IndexOf("(");
            string name = _firstCustomerFrom.Text.Substring(0, code).Trim();
            WaitForLoad();
            return name;
        }

        public string GetFirstCustomerFromCode()
        {
            WaitForLoad();
            var firstCustomerFrom = WaitForElementIsVisible(By.XPath(CUSTOMER_FROM));
            int startCode = firstCustomerFrom.Text.IndexOf("(");
            int endCode = firstCustomerFrom.Text.IndexOf(")");
            string code = firstCustomerFrom.Text.Substring(startCode + 1, endCode - (startCode + 1)).Trim();
            WaitForLoad();
            return code;
        }

        public string GetFirstCustomerToName()
        {
            _firstCustomerTo = WaitForElementIsVisible(By.XPath(CUSTOMER_TO));
            var code = _firstCustomerTo.Text.IndexOf("(");
            return _firstCustomerTo.Text.Substring(0, code).Trim();
        }

        public string GetFirstCustomerToCode()
        {
            _firstCustomerTo = WaitForElementIsVisible(By.XPath(CUSTOMER_TO));
            int startCode = _firstCustomerTo.Text.IndexOf("(");
            int endCode = _firstCustomerTo.Text.IndexOf(")");

            return _firstCustomerTo.Text.Substring(startCode + 1, endCode - (startCode + 1)).Trim();
        }
    }
}