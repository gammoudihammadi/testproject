using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Collections.Generic;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs
{
    public class ScheduledJobsPage : PageBase
    {
        public ScheduledJobsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        //-----------Table Liste------------
        private const string ROWS_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr[*]";
        private const string RESULTS_ROWS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]";
        //___________________________________ Constantes _____________________________________________

        //----------Nav Items ----------------
        private const string NAV_TO_BANK = "//*[@id=\"nav-tab\"]/li[3]/a";
        private const string NAV_TO_FILE_FLOWS = "//*[@id=\"nav-tab\"]/li[1]/a";
        private const string NAV_TO_CEGID = "//*[@id=\"nav-tab\"]/li[2]/a";
        //--------- ELEMENTS -----------------
        private const string RUN_JOB_BUTTON = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[11]/a[1]";
        private const string CONFIRM_BUTTON = "dataConfirmOK";
        private const string JOB_STATUS_CARD = "//*[@id=\"div-body\"]/div/div[1]/div[2]";
        private const string FIRST_JOB_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr[1]";
        private const string SHOW_EDITOR_BUTTON = "btnShowEditor";
        private const string SECOND_TAB = "//*[@id=\"crongenerator\"]/li[1]/a";
        private const string MINUTE_TAB = "//*[@id=\"crongenerator\"]/li[2]/a";
        private const string HOUR_TAB = "//*[@id=\"crongenerator\"]/li[3]/a";
        private const string DAY_TAB = "//*[@id=\"crongenerator\"]/li[4]/a";
        private const string MONTH_TAB = "//*[@id=\"crongenerator\"]/li[5]/a";
        private const string YEAR_TAB = "//*[@id=\"crongenerator\"]/li[6]/a";
        private const string FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr[1]";
        private const string NAME_FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr/td[2]";
        private const string DELETE_FIRST = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[11]/a[2]";
        private const string DELETE_CONFIRM = "dataConfirmOK";
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[1]";
        private const string CANCEL_FOR_JOB = "dataConfirmCancel";
        private const string CLOSE_FOR_EDIT  = "//*[@id=\"item-edit-form\"]/div[1]/button/span";
        


        // ----------- Filters -------------
        private const string RESET_FILTER = "ResetFilter";
        private const string SEARCH_BY_NAME = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[2]/input";
        private const string SHOW_ALL = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[1]";
        private const string SHOW_ACTIVE_ONLY = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[2]";
        private const string SHOW_INACTIVE_ONLY = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[3]";
        private const string CUSTOMERS = "FileFlowScheduledJobFiltersSelectedCustomers_ms";
        private const string CONVERTER_TYPES = "FileFlowScheduledJobFiltersSelectedConverters_ms"; 
        private const string CUSTOMERS_SELECTED = "/html/body/div[10]/ul/li/label/input";
        private const string CONVERTER_TYPES_SELECTED = "/html/body/div[11]/ul/li/label/input";
        private const string SHOW_SELECTED = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input";
        private const string NAV_TO_SPECIFI = "//*[@id=\"nav-tab\"]/li[4]/a";


        private const string COL_CUSTOMER = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[6]";

        private const string COL_CONVERTERS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[10]";

        private const string PROVIDERS_NAMES = "//*[@id=\"tableListMenu\"]/thead/tr/th[2]";
        private const string CONVERTER_TYPE_FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[8]";
        private const string COMPANY_CODE_FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[7]";


        //_____________________________________ Variables _____________________________________________
        // General
        [FindsBy(How = How.XPath, Using = NAV_TO_BANK)]
        private IWebElement _nav_to_bank;
        [FindsBy(How = How.XPath, Using = RUN_JOB_BUTTON)]
        private IWebElement _runJobButton;
        [FindsBy(How = How.Id, Using = CONFIRM_BUTTON)]
        private IWebElement _confirmButton;
        [FindsBy(How = How.XPath, Using = JOB_STATUS_CARD)]
        private IWebElement _jobStatusCard;
        [FindsBy(How=How.XPath, Using = FIRST_JOB_ROW)]
        private IWebElement _firstJobRow;
        [FindsBy(How=How.Id, Using = SHOW_EDITOR_BUTTON)]
        private IWebElement _showEditorButton;
        
        [FindsBy(How = How.Id, Using = CANCEL_FOR_JOB)]
        private IWebElement _cancelForJob;

        [FindsBy(How = How.Id, Using = CLOSE_FOR_EDIT)]
        private IWebElement _closeForEdit; 

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = NAV_TO_FILE_FLOWS)]
        private IWebElement _nav_to_file_flows;

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_BY_NAME)]
        private IWebElement _searchByNameFilter;

        [FindsBy(How = How.XPath, Using = SHOW_ALL)]
        private IWebElement _showALLFilter;

        [FindsBy(How = How.XPath, Using = SHOW_ACTIVE_ONLY)]
        private IWebElement _showActiveOnlyFilter;

        [FindsBy(How = How.XPath, Using = SHOW_INACTIVE_ONLY)]
        private IWebElement _showInactiveOnlyFilter;

        [FindsBy(How = How.Id, Using = CUSTOMERS)]
        private IWebElement _customersFilter;

        [FindsBy(How = How.Id, Using = CONVERTER_TYPES)]
        private IWebElement _converterTypesFilter;

        [FindsBy(How = How.XPath, Using = NAV_TO_CEGID)]
        private IWebElement _cegidTab;

        [FindsBy(How = How.XPath, Using = NAV_TO_SPECIFI)]
        private IWebElement _nav_to_specific;

        [FindsBy(How = How.XPath, Using = DELETE_FIRST)]
        private IWebElement _deleteFirst;

        [FindsBy(How = How.Id, Using = DELETE_CONFIRM)]
        private IWebElement _confirmDelete;

        public void RunJob()
        {
            WaitPageLoading();
            _runJobButton = WaitForElementIsVisible(By.XPath(RUN_JOB_BUTTON));
            _runJobButton.Click();
            WaitForLoad();
        }
        public void ConfirmRunJob()
        {
            _confirmButton = WaitForElementIsVisible(By.Id(CONFIRM_BUTTON));
            _confirmButton.Click();
            WaitForLoad();

        }
        public bool CheckOpenedConfirmRunJobModal()
        {
            WaitForLoad();
            return isElementExists(By.Id(CONFIRM_BUTTON));
        }
        public bool CheckOpenedStatusJobPage()
        {
            WaitForLoad();
            return isElementExists(By.XPath(JOB_STATUS_CARD));
        }
        public ScheduledJobsBankPage Nav_To_Bank()
        {

            _nav_to_bank = WaitForElementIsVisible(By.XPath(NAV_TO_BANK));
            _nav_to_bank.Click();
            WaitPageLoading();

            return new ScheduledJobsBankPage(_webDriver, _testContext);
        }


        public ScheduledJobsTabFileFlowsPage Nav_To_File_Flows()
        {

            _nav_to_file_flows = WaitForElementIsVisible(By.XPath(NAV_TO_FILE_FLOWS));
            _nav_to_file_flows.Click();
            WaitPageLoading();

            return new ScheduledJobsTabFileFlowsPage(_webDriver, _testContext);
        }


        public int GetNumberItemsDisplayed()
        {
            var items = _webDriver.FindElements(By.XPath(ROWS_NUMBER));
            return items.Count();
        }


        public enum FilterType
        {
            SearchByName,
            ShowAll,
            ShowActiveOnly,
            ShowInactiveOnly,
            Customer,
            ConverterTypes

        }

        public void Filter(FilterType FilterType, object value)
        {
            switch (FilterType)
            {
                case FilterType.SearchByName:
                    _searchByNameFilter = WaitForElementIsVisible(By.XPath(SEARCH_BY_NAME));
                    _searchByNameFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    WaitPageLoading();
                    break;
                case FilterType.Customer:
                    ComboBoxSelectById(new ComboBoxOptions(CUSTOMERS, (string)value));
                    WaitPageLoading();
                    break;
                case FilterType.ConverterTypes:
                    ComboBoxSelectById(new ComboBoxOptions(CONVERTER_TYPES, (string)value));
                    WaitPageLoading();
                    break;
                case FilterType.ShowAll:
                    _showALLFilter = WaitForElementExists(By.XPath(SHOW_ALL));
                    _showALLFilter.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;
                case FilterType.ShowActiveOnly:
                    _showActiveOnlyFilter = WaitForElementExists(By.XPath(SHOW_ACTIVE_ONLY));
                    _showActiveOnlyFilter.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;
                case FilterType.ShowInactiveOnly:
                    _showInactiveOnlyFilter = WaitForElementExists(By.XPath(SHOW_INACTIVE_ONLY));
                    _showInactiveOnlyFilter.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;
            }
            WaitForLoad();
        }

        public bool CheckAllRowsIsSpecificCustomer(string Customer)
        {

            var elements = _webDriver.FindElements(By.XPath(COL_CUSTOMER));
            foreach (var element in elements)
            {
                if (element.Text != Customer)
                {
                    return false;
                }
            }
            return true;
        }

        public bool CheckAllRowsIsSpecificConverters(string Converter)
        {
            var elements = _webDriver.FindElements(By.XPath(COL_CONVERTERS));
            foreach (var element in elements)
            {
                if (element.Text != Converter)
                {
                    return false;
                }
            }
            return true;
        }

        public ScheduledJobsCegidPage Nav_To_Cegid()
        {

            _cegidTab = WaitForElementIsVisible(By.XPath(NAV_TO_CEGID));
            _cegidTab.Click();
            WaitPageLoading();

            return new ScheduledJobsCegidPage(_webDriver, _testContext);
        }
        public void ClickOnFileFlowItem()
        {
            _firstJobRow = WaitForElementIsVisible(By.XPath(FIRST_JOB_ROW));
            _firstJobRow.Click();
            WaitForLoad();
        }
        public void ClickOnEditorButton()
        {
            _showEditorButton = WaitForElementIsVisible(By.Id(SHOW_EDITOR_BUTTON));
            _showEditorButton.Click();
            WaitForLoad();
        }
        public bool CheckTabsExistance()
        {
            return isElementExists(By.XPath(SECOND_TAB)) && isElementExists(By.XPath(MINUTE_TAB))
                && isElementExists(By.XPath(HOUR_TAB)) && isElementExists(By.XPath(DAY_TAB))
                && isElementExists(By.XPath(MONTH_TAB)) && isElementExists(By.XPath(YEAR_TAB));
        }

        public ScheduledSpecificPage Nav_To_Specific()
        {

            _nav_to_specific = WaitForElementIsVisible(By.XPath(NAV_TO_SPECIFI));
            _nav_to_specific.Click();
            WaitPageLoading();

            return new ScheduledSpecificPage(_webDriver, _testContext);
        }
        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
        }
        public FileFlowsEditModal SelectFirstRow()
        {
            WaitPageLoading();
            var firstRow = WaitForElementIsVisible(By.XPath(FIRST_ROW));
            firstRow.Click();
            WaitPageLoading();
            return new FileFlowsEditModal(_webDriver, _testContext);
        }
        public string GetNameFirstRow()
        {
            var nameRow = WaitForElementIsVisible(By.XPath(NAME_FIRST_ROW));
            return nameRow.Text.Trim();
        }

        public string GetConverterTypeFirstRow()
        {
            var converter = WaitForElementIsVisible(By.XPath(CONVERTER_TYPE_FIRST_ROW));
            return converter.Text.Trim();
        }

        public string GetCompanyCodeFirstRow()
        {
            var companyCode = WaitForElementIsVisible(By.XPath(COMPANY_CODE_FIRST_ROW));
            return companyCode.Text.Trim();
        }
        public List<string> GetNameProvidersList()
        {
            var NameProvidersListId = new List<string>();

            var NameProvidersList = _webDriver.FindElements(By.XPath(PROVIDERS_NAMES));

            foreach (var fileFlowProviders in NameProvidersList)
            {
                NameProvidersListId.Add(fileFlowProviders.Text);
            }

            return NameProvidersListId;
        }
        public int GetTotalResultRowsPage()
        {
            var rowsResults = _webDriver.FindElements(By.XPath(RESULTS_ROWS));
            WaitForLoad();
            return rowsResults.Count;
        }
        public void DeleteFirstFileFlow()
        {
            WaitPageLoading(); 
            _deleteFirst = WaitForElementIsVisible(By.XPath(DELETE_FIRST));
            _deleteFirst.Click();

            _confirmDelete = WaitForElementIsVisible(By.Id(DELETE_CONFIRM));
            _confirmDelete.Click();
            WaitPageLoading();
        }
        public void BackToList()
        {
            WaitPageLoading();
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();

        }
        public void CancelForJob()
        {
            WaitPageLoading();
            _cancelForJob = WaitForElementIsVisible(By.Id(CANCEL_FOR_JOB));
            _cancelForJob.Click();

        }
        public void CloseForEdit()
        {
            WaitPageLoading();
            _closeForEdit = WaitForElementIsVisible(By.XPath(CLOSE_FOR_EDIT)); 
            _closeForEdit.Click();

        }
        public void Nav_To_Scheduled()
        {
            var x = WaitForElementIsVisible(By.Id("ScheduledJobsTabNav"));
            x.Click();
            WaitForLoad();
        }
    }
}
