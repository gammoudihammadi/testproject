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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerPage : PageBase
    {

        public CustomerPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //___________________________________ Constantes _____________________________________________

        // General
        private const string NEW_CUSTOMER = "New customer";
        private const string EXPORT = "btn-export-excel";

        // Tableau
        private const string FIRST_CUSTOMER_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[1]/table/tbody/tr/td[3]";
        private const string FIRST_CUSTOMER = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[1]";
        private const string ALL_CUSTOMERS_NAME = "//*[@id=\"list-item-with-action\"]/div/div[*]/div/div/div[1]/table/tbody/tr/td[3]";
        private const string FIRST_CUSTOMER_ICAO = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[1]/table/tbody/tr/td[2]";
        
        // Filtres
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string RESET_FILTER_PATCH = "//*[@id=\"formSearchCustomers\"]/div[1]/a";

        private const string SEARCH_FILTER = "tbSearchPattern";
        private const string SORTBY_FILTER = "cbSortBy";

        private const string SHOW_ALL_FILTER_DEV = "ShowAll";
        private const string SHOW_ALL_FILTER_PATCH = "//*[@id=\"ShowActive\"][@value=\"All\"]";

        private const string SHOW_ACTIVE_FILTER_DEV = "ShowOnlyActive";
        private const string SHOW_ACTIVE_FILTER_PATCH = "//*[@id=\"ShowActive\"][@value=\"ActiveOnly\"]";

        private const string SHOW_INACTIVE_FILTER_DEV = "ShowOnlyInactive";
        private const string SHOW_INACTIVE_FILTER_PATCH = "//*[@id=\"ShowActive\"][@value=\"InactiveOnly\"]";

        private const string TYPE_OF_CUSTOMER = "//*[@id=\"collapseCustomerTypesFilter\"]/label[text()='{0}']";

        private const string CUSTOMER_TYPE_FILTER = "SelectedCustomerTypes_ms";
        private const string CUSTOMER_TYPE_UNSELECT_ALL = "/html/body/div[10]/div/ul/li[2]/a";
        private const string CUSTOMER_TYPE_SELECT_ALL = "/html/body/div[10]/div/ul/li[1]/a/span[2]";
        private const string SEARCH_CUSTOMER_TYPE = "/html/body/div[10]/div/div/label/input";

        private const string INACTIVE = "//*[@id=\"list-item-with-action\"]/div/div[{0}]/div/div/div[1]/table/tbody/tr/td[1]/img[@alt='Inactive']";
        private const string CUSTOMER_NAME = "//*[@id=\"list-item-with-action\"]/div/div[*]/div/div/div[1]/table/tbody/tr/td[3]";
        private const string CUSTOMER_TYPE = "//*[@id=\"list-item-with-action\"]/div/div[*]/div/div/div[1]/table/tbody/tr/td[4]";
        private const string CUSTOMER_NAME_CREATE_ID = "Customer_Name";
        private const string CUSTOMER_CODE_CREATE_ID = "Customer_Code";
        private const string CUSTOMER_SIRET_CREATE_ID = "Customer_SIRET";
        private const string CREATE_CUSTOMER = "Edit-Btn";

        //_____________________________________ Variables _____________________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_CUSTOMER)]
        private IWebElement _createNewCustomer;

        [FindsBy(How = How.XPath, Using = EXPORT)]
        private IWebElement _export;

        // Tableau
        [FindsBy(How = How.XPath, Using = FIRST_CUSTOMER_NAME)]
        private IWebElement _firstCustomerName;

        [FindsBy(How = How.XPath, Using = FIRST_CUSTOMER_ICAO)]
        private IWebElement _firstCustomerIcao;

        //______________________________________ Filtres _______________________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;
        
        [FindsBy(How = How.XPath, Using = RESET_FILTER_PATCH)]
        private IWebElement _resetFilterPatch;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortByBtn;

        [FindsBy(How = How.Id, Using = SHOW_ALL_FILTER_DEV)]
        private IWebElement _showAllDev;
        
        [FindsBy(How = How.XPath, Using = SHOW_ALL_FILTER_PATCH)]
        private IWebElement _showAllPatch;

        [FindsBy(How = How.Id, Using = SHOW_ACTIVE_FILTER_DEV)]
        private IWebElement _showActiveDev;
        
        [FindsBy(How = How.XPath, Using = SHOW_ACTIVE_FILTER_PATCH)]
        private IWebElement _showActivePatch;

        [FindsBy(How = How.Id, Using = SHOW_INACTIVE_FILTER_DEV)]
        private IWebElement _showInactiveDev;
        
        [FindsBy(How = How.XPath, Using = SHOW_INACTIVE_FILTER_PATCH)]
        private IWebElement _showInactivePatch;

        [FindsBy(How = How.XPath, Using = TYPE_OF_CUSTOMER)]
        private IWebElement _typeOfCustomer;

        [FindsBy(How = How.Id, Using = CUSTOMER_TYPE_FILTER)]
        private IWebElement _customerTypeFilter;

        [FindsBy(How = How.XPath, Using = CUSTOMER_TYPE_UNSELECT_ALL)]
        private IWebElement _customerTypeUnselectAll;

        [FindsBy(How = How.XPath, Using = CUSTOMER_TYPE_SELECT_ALL)]
        private IWebElement _customerTypeSelectAll;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER_TYPE)]
        private IWebElement _searchCustomerType;

        [FindsBy(How = How.Id, Using = CUSTOMER_NAME_CREATE_ID)]
        private IWebElement _customerName;

        [FindsBy(How = How.Id, Using = CUSTOMER_CODE_CREATE_ID)]
        private IWebElement _customerCode;

        [FindsBy(How = How.Id, Using = CUSTOMER_SIRET_CREATE_ID)]
        private IWebElement _customerSiret;  
        [FindsBy(How = How.Id, Using = FIRST_CUSTOMER)]
        private IWebElement _firstCustomer;

        public enum FilterType
        {
            Search,
            SortBy,
            ShowAll,
            ShowActive,
            ShowInactive,
            TypeOfCustomer
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
                    _sortByBtn = WaitForElementIsVisible(By.Id(SORTBY_FILTER));
                    _sortByBtn.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortByBtn.SetValue(ControlType.DropDownList, element.Text);
                    _sortByBtn.Click();
                    break;
                case FilterType.ShowAll:
                    if (isElementVisible(By.Id(SHOW_ALL_FILTER_DEV)))
                    {
                        _showAllDev = WaitForElementExists(By.Id(SHOW_ALL_FILTER_DEV));
                        _showAllDev.SetValue(ControlType.RadioButton, value);
                    }
                    else
                    {
                        _showAllPatch = WaitForElementExists(By.XPath(SHOW_ALL_FILTER_PATCH));
                        _showAllPatch.SetValue(ControlType.RadioButton, value);
                    }
                    
                    break;
                case FilterType.ShowActive:
                    if (isElementVisible(By.Id(SHOW_ACTIVE_FILTER_DEV)))
                    {
                        _showActiveDev = WaitForElementExists(By.Id(SHOW_ACTIVE_FILTER_DEV));
                        _showActiveDev.SetValue(ControlType.RadioButton, value);
                    }
                    else
                    {
                        _showActivePatch = WaitForElementExists(By.XPath(SHOW_ACTIVE_FILTER_PATCH));
                        _showActivePatch.SetValue(ControlType.RadioButton, value);
                    }
                    
                    break;
                case FilterType.ShowInactive:
                    if (isElementVisible(By.Id(SHOW_INACTIVE_FILTER_DEV)))
                    {
                        _showInactiveDev = WaitForElementExists(By.Id(SHOW_INACTIVE_FILTER_DEV));
                        _showInactiveDev.SetValue(ControlType.RadioButton, value);
                    }
                    else
                    {
                        _showInactivePatch = WaitForElementExists(By.XPath(SHOW_INACTIVE_FILTER_PATCH));
                        _showInactivePatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.TypeOfCustomer:
                    if (isElementVisible(By.Id(CUSTOMER_TYPE_FILTER)))
                    {
                        _customerTypeFilter = WaitForElementIsVisible(By.Id(CUSTOMER_TYPE_FILTER));
                        _customerTypeFilter.Click();

                        if (value.Equals("ALL CUSTOMER TYPES"))
                        {
                            _searchCustomerType = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER_TYPE));
                            _searchCustomerType.SetValue(ControlType.TextBox, "");

                            _customerTypeSelectAll = WaitForElementIsVisible(By.XPath(CUSTOMER_TYPE_SELECT_ALL));
                            _customerTypeSelectAll.Click();
                        }
                        else
                        {
                            _customerTypeUnselectAll = WaitForElementIsVisible(By.XPath(CUSTOMER_TYPE_UNSELECT_ALL));
                            _customerTypeUnselectAll.Click();

                            _searchCustomerType = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER_TYPE));
                            _searchCustomerType.SetValue(ControlType.TextBox, value);

                            if (isElementVisible(By.XPath("//span[text()='" + value + "']")))
                            {
                                var valueToCheckCustomersType = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                                valueToCheckCustomersType.SetValue(ControlType.CheckBox, true);
                            }
                        }

                        _customerTypeFilter.Click();
                    }
                    else
                    {
                        _typeOfCustomer = WaitForElementIsVisible(By.XPath(string.Format(TYPE_OF_CUSTOMER, value)));
                        _typeOfCustomer.Click();
                    }

                    break;
                default:
                    break;
            }

            WaitPageLoading();
            WaitForLoad();
        }


        public void ResetFilters()
        {
            if (isElementVisible(By.Id(RESET_FILTER_DEV)))
            {
                _resetFilterDev = WaitForElementIsVisible(By.Id(RESET_FILTER_DEV));
                _resetFilterDev.Click();
            }
            else
            {
                _resetFilterPatch = WaitForElementIsVisible(By.XPath(RESET_FILTER_PATCH));
                _resetFilterPatch.Click();
            }
            
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public bool IsSortedByName()
        {
            var listeCustomers = _webDriver.FindElements(By.XPath(CUSTOMER_NAME));

            if (listeCustomers.Count == 0)
                return false;

            var ancientCustomer = listeCustomers[0].Text;

            foreach (var elm in listeCustomers)
            {
                try
                {
                    if (elm.Text.CompareTo(ancientCustomer) < 0)
                        return false;

                    ancientCustomer = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsSortedByCustomerType()
        {
            var listeCustomers = _webDriver.FindElements(By.XPath(CUSTOMER_TYPE));

            if (listeCustomers.Count == 0)
                return false;

            var ancientCustomer = listeCustomers[0].Text;

            foreach (var elm in listeCustomers)
            {
                try
                {
                    if (elm.Text.CompareTo(ancientCustomer) < 0)
                        return false;

                    ancientCustomer = elm.Text;
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
            bool isActive = false;
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
                    isActive = true;
                    if (!active)
                        return true;
                }
            }
            return isActive;
        }

        public bool VerifyTypeCustomer(string value)
        {
            var listeCustomers = _webDriver.FindElements(By.XPath(CUSTOMER_TYPE));

            if (listeCustomers.Count == 0)
                return false;

            foreach (var elm in listeCustomers)
            {
                if (elm.Text != value)
                {
                    return false;
                }
            }

            return true;
        }

        //_______________________________________ Methodes ____________________________________________

        // General
        public CustomerCreateModalPage CustomerCreatePage()
        {
            ShowPlusMenu();

            _createNewCustomer = WaitForElementIsVisible(By.LinkText(NEW_CUSTOMER));
            _createNewCustomer.Click();
            WaitForLoad();

            return new CustomerCreateModalPage(_webDriver, _testContext);
        }
        public CustomerCreateModalPage CustomerCreate()
        {
            MakeSureCreateButtonIsVisible(3);

            var createNewCustomer = WaitForElementIsVisible(By.Id(CREATE_CUSTOMER));
            createNewCustomer.Click();
            WaitForLoad();

            return new CustomerCreateModalPage(_webDriver, _testContext);
        }
        public void MakeSureCreateButtonIsVisible(int tries)
        {
            for (int i = 0; i < tries; i++)
            {
                if (!isElementVisible(By.Id(CREATE_CUSTOMER)))
                {
                    ShowPlusMenu();
                }
                else
                {
                    return;
                }
            }
            
        }
        public void ExportCustomer(bool printVersion)
        {
            // Export du fichier
            ShowExtendedMenu();

            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            if (printVersion)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles, string customerName = null)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsExcelFileCorrect(file.Name, customerName, false))
                {
                    correctDownloadFiles.Add(file);
                } else if (IsExcelFileCorrect(file.Name, customerName, true)) {
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

        public bool IsExcelFileCorrect(string filePath, string customerName = null, bool ancienNomFichier = false)
        {
            Match match;
            if (ancienNomFichier)
            {
                if (customerName != null)
                {
                    // export-customers-contacts-ConsumerPriceIndexes 2 EXCEL AVIATION LTD-.xlsx
                    Regex r = new Regex("^export-customers-contacts " + customerName + "-.*" + "\\.xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    match = r.Match(filePath);
                }
                else
                {
                    //export-customers-contacts -.xlsx
                    Regex r = new Regex("^export-customers-contacts[ \\s]+-.xlsx", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    match = r.Match(filePath);
                }

            }
            else
            {
                if (customerName != null)
                {
                    // export-customers-contacts 2 EXCEL AVIATION LTD-.xlsx
                    Regex r = new Regex("^export-customers-contacts-ConsumerPriceIndexes " + customerName + "-.*" + "\\.xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    match = r.Match(filePath);
                }
                else
                {
                    // export-customers-contacts 2 EXCEL AVIATION LTD-.xlsx
                    Regex r = new Regex("^export-customers-contacts-ConsumerPriceIndexes[ \\s]+-.xlsx", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    match = r.Match(filePath);
                }

            }

            return match.Success;
        }

        // Tableau
        public CustomerGeneralInformationPage SelectFirstCustomer()
        {
            WaitPageLoading();
            _firstCustomer = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER));
            _firstCustomer.Click();
            WaitPageLoading();

            return new CustomerGeneralInformationPage(_webDriver, _testContext);
        }

        public string GetFirstCustomerName()
        {
            _firstCustomerName = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER_NAME));
            return _firstCustomerName.Text;
        }
        
        public string GetFirstCustomerIcao()
        {
            _firstCustomerIcao = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER_ICAO));
            return _firstCustomerIcao.Text;
        }
        public void CreateCustomerChorus(string code)
        {
            _customerName = WaitForElementIsVisible(By.Id(CUSTOMER_NAME_CREATE_ID));
            _customerName.SetValue(ControlType.TextBox, "Test Cust" + code);
            _customerCode = WaitForElementIsVisible(By.Id(CUSTOMER_CODE_CREATE_ID));
            _customerCode.SetValue(ControlType.TextBox, code);
            _customerSiret = WaitForElementIsVisible(By.Id(CUSTOMER_SIRET_CREATE_ID));
            _customerSiret.SetValue(ControlType.TextBox, "12345678912345");
            var chorusCheckBox = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div[1]/div/div[3]/div[2]/div[1]/div"));
            chorusCheckBox.Click();
            var saveButton = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div[2]/div/button[2]"));
            saveButton.Click();
        }
        public string GetColectividadesCustomer()
        {
            Filter(FilterType.Search, "A.F");
            Filter(FilterType.TypeOfCustomer, "Colectividades");
            WaitPageLoading();
           var TotalColectividades=  CheckTotalNumber();
            if (TotalColectividades > 0)
            {
                return GetFirstCustomerName();
            }
            return "";
        }
      
    }
}
