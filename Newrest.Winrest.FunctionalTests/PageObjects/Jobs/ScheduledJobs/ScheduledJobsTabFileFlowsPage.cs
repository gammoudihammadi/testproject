using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.Utils;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using DocumentFormat.OpenXml.Drawing;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs
{
    public class ScheduledJobsTabFileFlowsPage : PageBase
    {
        private const string PLUS_BTN = "//*[@id=\"nav-tab\"]/li[5]/div/button";
        private const string NEW_FILE_FLOW = "//*[@id=\"nav-tab\"]/li[5]/div/div/div/a";
        private const string CUSTOMERS_FILTER = "FileFlowScheduledJobFiltersSelectedCustomers_ms";
        private const string CONVERTER_FILTER = "FileFlowScheduledJobFiltersSelectedConverters_ms";
        private const string CONVERTER_FILTER_SELECTED = "/html/body/div[11]/ul/li/label/input";
        private const string FILE_FLOWS = "//*[@id=\"nav-tab\"]/li[1]/a";
        private const string CUSTOMERS_XPATH = "//*[@id=\"tableListMenu\"]/tbody/tr/td[2]";
        private const string UNSELECT_ALL_CUSTOMERS = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SELECT_ALL_CUSTOMERS = "/html/body/div[10]/div/ul/li[1]/a/span[2]";
        private const string CUSTOMERS_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string SHOW_ALL = "ShowAll";
        private const string SHOW_ACTIVE_ONLY = "ShowOnlyActive";
        private const string SHOW_INACTIVE_ONLY = "ShowOnlyInactive";
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string UNSELECT_ALL_CONVERTER_TYPE = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string SELECT_ALL_CONVERTER_TYPE = "/html/body/div[11]/div/ul/li[1]/a/span[2]";
        private const string CONVERTER_SEARCH = "/html/body/div[11]/div/div/label/input";
        private const string DELETE_FIRST = "//*[@id=\"tableListMenu\"]/tbody/tr/td[11]/a[2]";
        private const string DELETE_CONFIRM = "dataConfirmOK";
        private const string DELETE_CANCEL = "dataConfirmCancel";
        private const string RESULTS_ROWS = "//*[@id=\"tableListMenu\"]/tbody/tr";
        private const string NAME_FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr/td[2]";
        private const string FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr";
        private const string SHOW_SELECTED = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[@checked]";
        private const string CUSTOMERS_SELECTED = "/html/body/div[10]/ul/li/label/input";
        private const string ROWS_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr[*]";
        private const string FIRST_ROW_FREQUENCY = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[3]";
        private const string CONVERTER_FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr/td[10]";
        private const string CUSTOMER_FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr/td[6]";

        private const string LIST_ELEMENTS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[2]";


        // Buttons Modal Create
        private const string CLOSE_BTN = "//*[@id=\"item-edit-form\"]/div[5]/button[1]";
        private const string COL_IS_ACTIVE = "//*[@id=\"item_IsActive\"]";
        private const string SEARCH_BY_NAME = "SearchPattern";
        private const string CLICK_POUBELLE = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[11]/a[2]";
        

        private const string FIRST_JOB_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr[1]";

        private const string CRAETE_BTN = "last";
        private const string NAME = "//*[@id=\"Name\"]";
        private const string DROPDOWN = "//*[@id=\"drop-down-customers\"]";
        private const string FIRSTLINE = "//*[@id=\"drop-down-customers\"]/option[2]";
        private const string ACTIVATED = "//*[@id=\"item-edit-form\"]/div[2]/div[8]/div/div";
        private const string LIST_ROWS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]";

        [FindsBy(How = How.XPath, Using = FIRST_JOB_ROW)]
        private IWebElement _firstJobRow;
        [FindsBy(How = How.XPath, Using = ACTIVATED)]
        private IWebElement _activated;

        [FindsBy(How = How.XPath, Using = FIRSTLINE)]
        private IWebElement _firstline;

        [FindsBy(How = How.XPath, Using = DROPDOWN)]
        private IWebElement _dropdown;
        [FindsBy(How = How.XPath, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusbtn;

        [FindsBy(How = How.XPath, Using = NEW_FILE_FLOW)]
        private IWebElement _newFileFlow;

        [FindsBy(How = How.Id, Using = CUSTOMERS_FILTER)]
        private IWebElement _customersFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_CUSTOMERS)]
        private IWebElement _unselectAllCustomers;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_CUSTOMERS)]
        private IWebElement _selectAllCustomers;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_SEARCH)]
        private IWebElement _searchCustomers;

        [FindsBy(How = How.Id, Using = CONVERTER_FILTER)]
        private IWebElement _converterFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_CONVERTER_TYPE)]
        private IWebElement _unselectAllConverterType;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_CONVERTER_TYPE)]
        private IWebElement _selectAllConverterType;

        [FindsBy(How = How.XPath, Using = CONVERTER_SEARCH)]
        private IWebElement _searchConverterType;

        [FindsBy(How = How.XPath, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.Id, Using = SHOW_ACTIVE_ONLY)]
        private IWebElement _showActiveOnlyFilter;

        [FindsBy(How = How.Id, Using = SHOW_INACTIVE_ONLY)]
        private IWebElement _showInactiveOnlyFilter;

        [FindsBy(How = How.Id, Using = SHOW_ALL)]
        private IWebElement _showAll;

        [FindsBy(How = How.XPath, Using = DELETE_FIRST)]
        private IWebElement _deleteFirst;

        [FindsBy(How = How.Id, Using = DELETE_CONFIRM)]
        private IWebElement _confirmDelete;
        [FindsBy(How = How.XPath, Using = RESULTS_ROWS)]
        private IWebElement _rowsResults;
        // Buttons Modal Create 

        [FindsBy(How = How.XPath, Using = CLOSE_BTN)]
        private IWebElement _close_btn;

        [FindsBy(How = How.Id, Using = DELETE_CANCEL)]
        private IWebElement _clickpoubellecancel;

        [FindsBy(How = How.Id, Using = SEARCH_BY_NAME)]
        private IWebElement _searchByNameFilter;


        [FindsBy(How = How.XPath, Using = CLICK_POUBELLE)]
        private IWebElement _clickpoubelle;

        [FindsBy(How=How.XPath, Using = FIRST_ROW_FREQUENCY)]
        private IWebElement _firstRowFrequency;

        [FindsBy(How = How.Id, Using = CRAETE_BTN)]
        private IWebElement _createBtn;


        public int GetCountResultPaged()
        {
            var elements = _webDriver.FindElements(By.XPath(LIST_ROWS));
            return elements.Count;
        }
        public void Createfileflowshudlerjob(string name )
        {
            _name = WaitForElementIsVisible(By.XPath(NAME));
            _name.SetValue(ControlType.TextBox, name);
            _activated = WaitForElementIsVisible(By.XPath(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, true);
            _dropdown = WaitForElementIsVisible(By.XPath(DROPDOWN));
            _dropdown.Click();
            _firstline = WaitForElementIsVisible(By.XPath(FIRSTLINE));
            _firstline.Click();


        }
        public ScheduledJobsTabFileFlowsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public enum FilterType
        {
            Customers,
            Converter,
            ShowAll,
            ShowActiveOnly,
            ShowInactiveOnly,
            SearchByName,
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Customers:
                    ComboBoxSelectById(new ComboBoxOptions(CUSTOMERS_FILTER, (string)value));

                    break;

                case FilterType.Converter:
                    ComboBoxSelectById(new ComboBoxOptions(CONVERTER_FILTER, (string)value));

                    break;
                case FilterType.ShowActiveOnly:
                    _showActiveOnlyFilter = WaitForElementExists(By.Id(SHOW_ACTIVE_ONLY));
                    _showActiveOnlyFilter.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;
                case FilterType.ShowInactiveOnly:
                    _showInactiveOnlyFilter = WaitForElementExists(By.Id(SHOW_INACTIVE_ONLY));
                    _showInactiveOnlyFilter.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;
                case FilterType.ShowAll:
                    _showAll = WaitForElementIsVisible(By.Id(SHOW_ALL));
                    _showAll.SetValue(ControlType.RadioButton, true);
                    WaitForLoad();
                    break;
                case FilterType.SearchByName:
                    _searchByNameFilter = WaitForElementIsVisible(By.Id(SEARCH_BY_NAME));
                    _searchByNameFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    WaitPageLoading();
                    break;

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void ShowBtnPlus()
        {
            _plusbtn = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            _plusbtn.Click();
        }


        public FileFlowsCreateModalPage OpenModalCreateFileFLows()
        {
            ShowBtnPlus();
            _newFileFlow = WaitForElementIsVisible(By.XPath(NEW_FILE_FLOW));
            _newFileFlow.Click();


            return new FileFlowsCreateModalPage(_webDriver, _testContext);
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

        public void ResetFilter()
        {
            _resetFilterDev = WaitForElementIsVisible(By.Id(RESET_FILTER_DEV));
            _resetFilterDev.Click();
            WaitForLoad();

            if (DateUtils.IsBeforeMidnight())
            {
                //pas de FilterType.DateTo implémenté dans le switch case
                //pas de DateTo dans l'IHM
            }
        }
        public int GetTotalList()
        {
            var fileFlowsList = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]"));
            return fileFlowsList.Count;
        }
        public void DeleteFirstFileFlow()
        {
            _deleteFirst = WaitForElementIsVisible(By.XPath(DELETE_FIRST));
            _deleteFirst.Click();

            _confirmDelete = WaitForElementIsVisible(By.Id(DELETE_CONFIRM));
            _confirmDelete.Click();
        }
        public int GetTotalResultRowsPage()
        {
            var rowsResults = _webDriver.FindElements(By.XPath(RESULTS_ROWS));
            WaitForLoad();
            return rowsResults.Count;
        }

        public bool CheckAllRowsIsActiveOrInactive(bool IsActive)
        {

            var elements = _webDriver.FindElements(By.XPath(COL_IS_ACTIVE));
            foreach (var element in elements)
            {
                if (element.Selected != IsActive)
                {
                    return false;
                }
            }
            return true;
        }
        public int NombreRow()
        {
            var row = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr"));
            return row.Count;

        }
        public string GetNameFirstRow()
        {
            var nameRow = WaitForElementIsVisible(By.XPath(NAME_FIRST_ROW));
            return nameRow.Text.Trim();
        }

        public string GetFirstLineConverterType()
        {
            var converterRow = WaitForElementIsVisible(By.XPath(CONVERTER_FIRST_ROW));
            return converterRow.Text.Trim();
        }
        public string GetFirstLineCustomer()
        {
            var customerRow = WaitForElementIsVisible(By.XPath(CUSTOMER_FIRST_ROW));
            return customerRow.Text.Trim();
        }
        public string GetShowFilterSelected()
        {

            var itemSelected = _webDriver.FindElements(By.XPath(SHOW_SELECTED));
            var showName = string.Empty;
            foreach (var element in itemSelected)
            {
                if (element.Selected)
                {
                    //var x = element.GetAttribute("value");
                    switch (element.GetAttribute("value"))
                    {
                        case "All":
                            showName = "Show all";
                            break;
                        case "ActiveOnly":
                            showName = "Show active only";
                            break;
                        case "InactiveOnly":
                            showName = "Show inactive only";
                            break;


                    }
                }
            }

            return showName;
        }
        public List<string> GetSelectedCustomersToFilter()
        {
            var customersSelected = new List<string>();
            _customersFilter = WaitForElementIsVisible(By.Id(CUSTOMERS_FILTER));
            _customersFilter.Click();
            WaitForLoad();
            var selectElement = _webDriver.FindElements(By.XPath(CUSTOMERS_SELECTED));
            for (var i = 0; i < selectElement.Count; i++)
            {
                if (selectElement[i].Selected)
                {
                    var element = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[10]/ul/li[{0}]/label/input/../span", i + 1)));
                    customersSelected.Add(element.Text);
                }
            }

            _customersFilter.Click();
            return customersSelected;
        }
        public List<string> GetConverterSelectedToFilter()
        {
            var converterTypesSelected = new List<string>();
            _converterFilter = WaitForElementIsVisible(By.Id(CONVERTER_FILTER));
            _converterFilter.Click();
            WaitForLoad();
            var selectElement = _webDriver.FindElements(By.XPath(CONVERTER_FILTER_SELECTED));
            for (var i = 0; i < selectElement.Count; i++)
            {
                if (selectElement[i].Selected)
                {
                    var element = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[11]/ul/li[{0}]/label/input/../span", i + 1)));
                    converterTypesSelected.Add(element.Text);
                }
            }

            _converterFilter.Click();
            return converterTypesSelected;
        }
        public int GetNumberItemsDisplayed()
        {
            var items = _webDriver.FindElements(By.XPath(ROWS_NUMBER));
            return items.Count();
        }
        public void ClickSurPoubelle()
        {

            var _clickpoubelle = WaitForElementIsVisible(By.XPath(CLICK_POUBELLE));
            _clickpoubelle.Click();
            WaitForLoad();
        }
        public void DeleteConfirm()
        {
            _confirmDelete = WaitForElementIsVisible(By.Id(DELETE_CONFIRM));
            _confirmDelete.Click();
        }
        public void DeleteCancel()
        {
            _clickpoubellecancel = WaitForElementIsVisible(By.Id(DELETE_CANCEL));
            _clickpoubellecancel.Click();
        }
        public void ClickSurPoubelleDeleteConfirm(string namefordelete)
        {
            var DeleteButtons = WaitForElementIsVisible(By.XPath(String.Format(DELETE_FIRST, namefordelete)));
            DeleteButtons.Click();
            var deleteButton = _webDriver.FindElement(By.XPath("//*[@id=\"dataConfirmOK\"]"));
            deleteButton.Click();

        }
        public List<string> GetElementsNamesFileFlowProviders()
        {
            var ListElemets = _webDriver.FindElements(By.XPath(LIST_ELEMENTS));
            List<string> ListElemetsNames = new List<string>();
            //var ListElemetsNames =  new string[ListElemets.Count];
            foreach (var element in ListElemets)
            {
                ListElemetsNames.Add(element.Text);
            }
            return ListElemetsNames;
        }

        public bool CheckPopupFileFlowsDeleteIsVisible()
        {
            return isElementVisible(By.Id(DELETE_CONFIRM)) && isElementVisible(By.Id(DELETE_CANCEL));

        }
        public FileFlowsCreateModalPage EditFileFlow()
        {
            _firstJobRow = WaitForElementIsVisible(By.XPath(FIRST_JOB_ROW));
            _firstJobRow.Click();
            WaitForLoad();
            return new FileFlowsCreateModalPage(_webDriver, _testContext);
        }
        public string GetFirstRowFrequency()
        {
            _firstRowFrequency = WaitForElementExists(By.XPath(FIRST_ROW_FREQUENCY));
            return _firstRowFrequency.Text;
        }

        public void CreateBTN()
        {
            _createBtn = WaitForElementIsVisible(By.Id(CRAETE_BTN));
            _createBtn.Click();

        }
    }
}
