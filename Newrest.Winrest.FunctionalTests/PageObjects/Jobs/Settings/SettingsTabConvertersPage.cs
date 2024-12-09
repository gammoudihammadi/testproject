using DocumentFormat.OpenXml.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public class SettingsTabConvertersPage: PageBase
    {
        private const string PLUS_BTN = "//*[@id=\"nav-tab\"]/li[8]/div/button";
        private const string NEW_CONVERTER = "//*[@id=\"nav-tab\"]/li[8]/div/div/div/a";
        private const string CUSTOMERS_FILTER = "ConverterFiltersSelectedCustomers_ms";
        private const string UNSELECT_ALL_CUSTOMERS = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SELECT_ALL_CUSTOMERS = "/html/body/div[10]/div/ul/li[1]/a/span[2]";
        private const string CUSTOMERS_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string CUSTOMERS_XPATH = "//*[@id='tableListMenu']/tbody/tr[*]/td[2]";
        private const string RESET_FILTER_DEV = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string CONVERTER = "//*[@id=\"nav-tab\"]/li[3]/a"; 
        private const string FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr";

        private const string SEARCH_BY_CUSTOMER_VALUE = "//*[@id=\"ConverterFiltersSelectedCustomers_ms\"]/span[2]"; 
        private const string CONVERTER_TYPE_FILTER = "ConverterFiltersSelectedConverterTypes_ms";
        private const string UNSELECT_ALL_CONVERTER_TYPE = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string SELECT_ALL_CONVERTER_TYPE = "/html/body/div[11]/div/ul/li[1]/a/span[2]";
        private const string CONVERTER_TYPE_SEARCH = "/html/body/div[11]/div/div/label/input";
        private const string CONVERTER_TYPE_XPATH = "//*[starts-with(@id,\"converter\")]/td[4]";
        private const string JOB_NOTIFICATION = "//*[@id=\"nav-tab\"]/li[6]/a";
        private const string LABEL_CUSTOMER = "//label[@for='CustomerId']";
        private const string LABEL__FILE_EXTENSION = "//label[@for='FileExtension']";
        private const string LABEL_CONVERTER_TYPE = "//label[@for='ConverterType']";
        //private const string CLICK_FIRST_CONVERTERS = "//*[@id=\"converter115\"]/td[5]/a/span";
        private const string CLICK_FIRST_CONVERTERS = "//*[starts-with(@id, 'converter')]/td[5]/a/span";

        private const string DELETE_CONVERTERS = "deleteConverter";
        private const string CLICK_FIRST_ITEM = "/html/body/div[2]/div/div[2]/div[2]/div/div/div/table/tbody/tr[1]";
        

        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusbtn;

        [FindsBy(How = How.XPath, Using = NEW_CONVERTER)]
        private IWebElement _newconverter;

        [FindsBy(How = How.Id, Using = CUSTOMERS_FILTER)]
        private IWebElement _customersFilter;
        [FindsBy(How = How.Id, Using = DELETE_CONVERTERS)]
        private IWebElement _deleteButton;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_CUSTOMERS)]
        private IWebElement _unselectAllCustomers;

        [FindsBy(How = How.XPath, Using = CLICK_FIRST_CONVERTERS)]
        private IWebElement _clickFirstConverters;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_CUSTOMERS)]
        private IWebElement _selectAllCustomers;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_SEARCH)]
        private IWebElement _searchCustomers;
        [FindsBy(How = How.XPath, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;
        [FindsBy(How = How.XPath, Using = CONVERTER)]
        private IWebElement _converter;
        [FindsBy(How = How.XPath, Using = SEARCH_BY_CUSTOMER_VALUE)]
        private IWebElement _searchbynamevalue;

        [FindsBy(How = How.Id, Using = CONVERTER_TYPE_FILTER)]
        private IWebElement _converterTypeFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_CONVERTER_TYPE)]
        private IWebElement _unselectAllConverterType;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_CONVERTER_TYPE)]
        private IWebElement _selectAllConverterType;


        [FindsBy(How = How.XPath, Using = CONVERTER_TYPE_SEARCH)]
        private IWebElement _searchConverterType;

        [FindsBy(How = How.XPath, Using = JOB_NOTIFICATION)]
        private IWebElement _jobnotification;



        public SettingsTabConvertersPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
            
        }
        public enum FilterType
        {
            Customers,
            CustomersCheckAll,
            CustomersUncheck,
            ConverterType,
            ConverterTypeCheckAll,
            ConverterTypeUncheckAll,
          
        }
        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {

                case FilterType.CustomersUncheck:
                    _customersFilter = WaitForElementIsVisible(By.Id(CUSTOMERS_FILTER));
                    _customersFilter.Click();

                    _unselectAllCustomers = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_CUSTOMERS));
                    _unselectAllCustomers.Click();

                    _customersFilter.Click();
                    break;

                case FilterType.CustomersCheckAll:
                    _customersFilter = WaitForElementIsVisible(By.Id(CUSTOMERS_FILTER));
                    _customersFilter.Click();

                    var selectAllCustomers = WaitForElementIsVisible(By.XPath(SELECT_ALL_CUSTOMERS));
                    selectAllCustomers.Click();

                    _customersFilter.Click();
                    break;

                case FilterType.Customers:
                    _customersFilter = WaitForElementIsVisible(By.Id(CUSTOMERS_FILTER));
                    _customersFilter.Click();

                    _unselectAllCustomers = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_CUSTOMERS));
                    _unselectAllCustomers.Click();

                    _searchCustomers = WaitForElementIsVisible(By.XPath(CUSTOMERS_SEARCH));
                    _searchCustomers.SetValue(ControlType.TextBox, value);
                    WaitForLoad();

                    var customerToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    customerToCheck.SetValue(ControlType.CheckBox, true);

                    _customersFilter.Click();
                    break;

                case FilterType.ConverterTypeUncheckAll:
                    _converterTypeFilter = WaitForElementIsVisible(By.Id(CONVERTER_TYPE_FILTER));
                    _converterTypeFilter.Click();

                    _unselectAllConverterType = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_CONVERTER_TYPE));
                    _unselectAllConverterType.Click();

                    _converterTypeFilter.Click();
                    break;

                case FilterType.ConverterTypeCheckAll:
                    _converterTypeFilter = WaitForElementIsVisible(By.Id(CONVERTER_TYPE_FILTER));
                    _converterTypeFilter.Click();

                    var selectAllConverterType = WaitForElementIsVisible(By.XPath(SELECT_ALL_CONVERTER_TYPE));
                    selectAllConverterType.Click();

                    _converterTypeFilter.Click();
                    break;

                case FilterType.ConverterType:
                    _converterTypeFilter = WaitForElementIsVisible(By.Id(CONVERTER_TYPE_FILTER));
                    _converterTypeFilter.Click();

                    _unselectAllConverterType= WaitForElementIsVisible(By.XPath(UNSELECT_ALL_CONVERTER_TYPE));
                    _unselectAllConverterType.Click();

                    _searchConverterType = WaitForElementIsVisible(By.XPath(CONVERTER_TYPE_SEARCH));
                    _searchConverterType.SetValue(ControlType.TextBox, value);
                    WaitForLoad();

                    var converterTypeToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    converterTypeToCheck.SetValue(ControlType.CheckBox, true);

                    _converterTypeFilter.Click();
                    break;
            }

            WaitPageLoading();
            WaitForLoad();
        }
        public void UncheckAllCustomers()
        {
            _customersFilter = WaitForElementIsVisible(By.Id(CUSTOMERS_FILTER));
            _customersFilter.Click();

            _unselectAllCustomers = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_CUSTOMERS));
            _unselectAllCustomers.Click();

            _customersFilter.Click();

            WaitPageLoading();
            Thread.Sleep(2000);
        }

        public void ShowBtnPlus()
        {
            _plusbtn = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            _plusbtn.Click();
        }
        public ConvertersCreateModalPage AddNewConverter()
        {
            ShowBtnPlus();
            _newconverter = WaitForElementIsVisible(By.XPath(NEW_CONVERTER));
            _newconverter.Click();
            WaitForLoad();
            return new ConvertersCreateModalPage(_webDriver, _testContext);
        }

        public List<string> GetCustomersFiltred()
        {
            var customers = _webDriver.FindElements(By.XPath(CUSTOMERS_XPATH));

            if (customers.Count == 0)
                return new List<string>();

            var filteredCustomers = new List<string>();

            foreach (var elm in customers)
            {
              
                    filteredCustomers.Add(elm.Text.Trim());

            }

            return filteredCustomers;
        }
        public void UncheckAllConverterType()
        {
            _converterTypeFilter = WaitForElementIsVisible(By.Id(CONVERTER_TYPE_FILTER));
            _converterTypeFilter.Click();

            _unselectAllConverterType = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_CONVERTER_TYPE));
            _unselectAllConverterType.Click();
            _converterTypeFilter.Click();
            WaitPageLoading();

        }
        public List<string> GetConverterTypeFiltred()
        {
            var converterType = _webDriver.FindElements(By.XPath(CONVERTER_TYPE_XPATH));
            if (converterType.Count == 0)
                return new List<string>();
            var filteredConverterType = new List<string>();
            foreach (var elm in converterType)
            {
                filteredConverterType.Add(elm.Text.Trim());
            }
            return filteredConverterType;
        }

        public void ResetFilter()
        {
            _resetFilterDev = WaitForElementIsVisible(By.XPath(RESET_FILTER_DEV));
            _resetFilterDev.Click();
            WaitForLoad();

            if (DateUtils.IsBeforeMidnight())
            {
                //pas de FilterType.DateTo implémenté dans le switch case
                //pas de DateTo dans l'IHM
            }
        }

        public string GetConvertersTitle()
        {
            var pagetitle = WaitForElementIsVisible(By.XPath(CONVERTER));
            return pagetitle.Text;
        }
        public string GetInputSearchCustomersValue()
        {
            _searchbynamevalue = WaitForElementIsVisible(By.XPath(SEARCH_BY_CUSTOMER_VALUE));
            return _searchbynamevalue.Text;
        }

        public ConverterDetailsPage SelectFirstRow()
        {
            WaitForLoad();
            var firstRow = WaitForElementIsVisible(By.XPath(FIRST_ROW));
            firstRow.Click();
            WaitPageLoading();
            return new ConverterDetailsPage(_webDriver, _testContext);
        }
        public void deleteFirstConverters()
        {
            _clickFirstConverters = WaitForElementIsVisible(By.XPath(CLICK_FIRST_CONVERTERS));
            _clickFirstConverters.Click();
            WaitPageLoading();

            _deleteButton = WaitForElementIsVisible(By.Id(DELETE_CONVERTERS));
            _deleteButton.Click();
            WaitForLoad();
        }

        public void DeleteFirstRow()
        {
            if (IsDev())
            {
                var customerrowDelete = WaitForElementIsVisible(By.XPath("//*[starts-with(@id,'converter')]/td[5]/a"));
                customerrowDelete.Click();
                WaitForLoad();
                var deleteButton = WaitForElementIsVisible(By.Id("deleteConverter"));
                deleteButton.Click();
                WaitForLoad();
            }
            else
            {
                var customerrowDelete = SolveVisible("//*/a[contains(@class,'btn-converter-delete')]");
                customerrowDelete.Click();
                WaitForLoad();
                var deleteButton = WaitForElementIsVisible(By.Id("deleteConverter"));
                deleteButton.Click();
                //WaitPageLoading();
                Thread.Sleep(2000);
                WaitForLoad();
            }
            
        }
        public JobNotifications GoTo_JOB_NOTIFICATION()
        {
            WaitForLoad();
            _jobnotification = WaitForElementIsVisible(By.XPath(JOB_NOTIFICATION));
            _jobnotification.Click();
            WaitForLoad();
            return new JobNotifications(_webDriver, _testContext);
        }
        public bool VerifyAllLabelsHaveClass()
        {
            var labels = new Dictionary<string, IWebElement>
    {
        { "Customer", WaitForElementIsVisible(By.XPath(LABEL_CUSTOMER) )},
        { "FileExtension", WaitForElementIsVisible(By.XPath(LABEL__FILE_EXTENSION)) },
        { "ConverterType",WaitForElementIsVisible(By.XPath(LABEL_CONVERTER_TYPE)) }
    };

            string expectedClass = "align-right-and-center";

            foreach (var label in labels.Values)
            {
                string classAttribute = label.GetAttribute("class");
                if (!classAttribute.Contains(expectedClass))
                {
                    return false; // Si une classe ne contient pas "align-right-and-center", on retourne false
                }
            }
            return true; // Si toutes les classes contiennent "align-right-and-center", on retourne true
        }

        public ConverterDetailsPage SelectFirstItem()
        {
            WaitForLoad();
            var firstItem = WaitForElementIsVisible(By.XPath(CLICK_FIRST_ITEM));
            firstItem.Click();
            WaitPageLoading();
            return new ConverterDetailsPage(_webDriver, _testContext);
        }




    }
}