using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Collections.Generic;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs
{
    public class ScheduledJobsBankPage : PageBase
    {
        public ScheduledJobsBankPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        //------------------------------ Button ------------------------------------------
        private const string ADD_BUTTON = "/html/body/div[2]/div/div[2]/div[1]/nav/ul/li[5]/div/button";
        private const string NEW_BANK_SCHEDULED_JOB = "/html/body/div[2]/div/div[2]/div[1]/nav/ul/li[5]/div/div/div/a";
        // Buttons Modal Create
        private const string CREATE_BTN = "last";
        private const string CLOSE_BTN = "/html/body/div[3]/div/div/div/div/div/div/form/div[4]/button[1]";
        //-------------------------- Inputs Create Modal --------------

        private const string NAME = "Name";
        private const string CUSTOMER = "CustomerId";
        private const string CONVERTER_TYPE = "ConverterType";
        private const string COMPANY_CODE = "CompanyCode";
        private const string SFTP_CONNECTION_STRING = "SFTPConnectionString";
        private const string SFTP_FOLDER = "SFTPFolder";
        private const string CRON_EXPRESSION = "//*[@id=\"CronExpression\"]";
        private const string IS_ACTIVE = "IsActive";

        // Filter
        private const string RESET_FILTER = "ResetFilter";
        private const string SEARCH_BY_NAME = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[2]/input";
        private const string SHOW_ALL = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[1]";
        private const string SHOW_ACTIVE_ONLY = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[2]";
        private const string SHOW_INACTIVE_ONLY = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[3]";
        private const string CUSTOMERS = "BankScheduledJobFiltersSelectedCustomers_ms";
        private const string CONVERTER_TYPES = "BankScheduledJobFiltersSelectedConverterTypes_ms";
        private const string CUSTOMERS_SELECTED = "/html/body/div[10]/ul/li/label/input";
        private const string CONVERTER_TYPES_SELECTED = "/html/body/div[11]/ul/li/label/input";
        private const string SHOW_SELECTED = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input";

        //-----------Table Liste------------
        private const string ROWS_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr[*]";
        private const string COL_CONVERTER_TYPE = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[8]";
        private const string RESULTS_ROWS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]";
        private const string COL_CUSTOMERS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[6]";
        private const string COL_IS_ACTIVE = "//*[@id=\"item_IsActive\"]";
        private const string PROVIDERS_NAMES = "//*[@id=\"tableListMenu\"]/thead/tr/th[2]";
        private const string NAMES_COL = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[2]";
        private const string EDITOR = "btnShowEditor";
        private const string SECOND_TAB = "//*[@id=\"crongenerator\"]/li[1]/a";
        private const string MINUTE_TAB = "//*[@id=\"crongenerator\"]/li[2]/a";
        private const string HOUR_TAB = "//*[@id=\"crongenerator\"]/li[3]/a";
        private const string DAY_TAB = "//*[@id=\"crongenerator\"]/li[4]/a";
        private const string MONTH_TAB = "//*[@id=\"crongenerator\"]/li[5]/a";
        private const string YEAR_TAB = "//*[@id=\"crongenerator\"]/li[6]/a";
        private const string FIRST_ROW_FREQUENCY = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[3]";
        private const string FIRST_JOB_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr[1]";
        private const string FIRST_LINE_NAME = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[2]";
        private const string NAME_VALIDATOR = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[1]/div/span";
        private const string SFTP_CONNECTION_STRING_VALIDATOR = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[5]/div/span";
        private const string SFTP_FOLDER_VALIDATOR = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[6]/div/span";
        private const string FIRST_LINE_CUSTOMER = "//*[@id=\"tableListMenu\"]/tbody/tr/td[6]";
        private const string FIRST_LINE_COMPANY_CODE = "//*[@id=\"tableListMenu\"]/tbody/tr/td[7]";
        private const string FIRST_LINE_CONVERTER = "//*[@id=\"tableListMenu\"]/tbody/tr/td[8]";




        //_______________________________________________________Variables___________________________________________

        // General
        [FindsBy(How = How.XPath, Using = ADD_BUTTON)]
        private IWebElement _add_button;

        [FindsBy(How = How.XPath, Using = NEW_BANK_SCHEDULED_JOB)]
        private IWebElement _new_bank;

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.Id, Using = CONVERTER_TYPE)]
        private IWebElement _converter_type;

        [FindsBy(How = How.Id, Using = COMPANY_CODE)]
        private IWebElement _company_code;

        [FindsBy(How = How.Id, Using = SFTP_CONNECTION_STRING)]
        private IWebElement _connection_string;

        [FindsBy(How = How.Id, Using = SFTP_FOLDER)]
        private IWebElement _sftp_folder;

        [FindsBy(How = How.XPath, Using = CRON_EXPRESSION)]
        private IWebElement cron_expression;

        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _is_active;

        [FindsBy(How = How.XPath, Using = RESULTS_ROWS)]
        private IWebElement _rowsResults;

        // Buttons Modal Create 
        [FindsBy(How = How.Id, Using = CREATE_BTN)]
        private IWebElement _create_btn;

        [FindsBy(How = How.XPath, Using = CLOSE_BTN)]
        private IWebElement _close_btn;

        // Table List 

        [FindsBy(How = How.XPath, Using = ROWS_NUMBER)]
        private IWebElement _rowsNumber;

        // Filter

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
        [FindsBy(How = How.Id, Using = EDITOR)]
        private IWebElement _editor;
        [FindsBy(How = How.XPath, Using = FIRST_ROW_FREQUENCY)]
        private IWebElement _firstRowFrequency;
        [FindsBy(How = How.XPath, Using = FIRST_JOB_ROW)]
        private IWebElement _firstJobRow;

        [FindsBy(How = How.XPath, Using = FIRST_LINE_NAME)]
        private IWebElement _firstLineName;

        [FindsBy(How = How.XPath, Using = FIRST_LINE_CUSTOMER)]
        private IWebElement _firstLineCustomer;

        [FindsBy(How = How.XPath, Using = FIRST_LINE_COMPANY_CODE)]
        private IWebElement _firstLineCompanyCode;

        [FindsBy(How = How.XPath, Using = FIRST_LINE_CONVERTER)]
        private IWebElement _firstLineConverter;

        //-----Filter------
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
        //________________________Pages________________________________

        public void OpenModalAddNewBank()
        {
            _add_button = WaitForElementIsVisible(By.XPath(ADD_BUTTON));
            _add_button.Click();
             _new_bank = WaitForElementIsVisible(By.XPath(NEW_BANK_SCHEDULED_JOB));
            _new_bank.Click();
         }
        public void AddNewBank(
            string name,
            string companyCode,
            string sftpConnectionString,
            string sftpFolder,
            bool isActive,
            string customer = null,
            string converterType = null)
        {
            OpenModalAddNewBank();

            SetName(name);

            _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SetValue(PageBase.ControlType.DropDownList, customer);
            WaitForLoad();

            _converter_type = WaitForElementIsVisible(By.Id(CONVERTER_TYPE));
            _converter_type.SetValue(PageBase.ControlType.DropDownList, converterType);
            WaitForLoad();

            _company_code = WaitForElementIsVisible(By.Id(COMPANY_CODE));
            _company_code.SetValue(PageBase.ControlType.TextBox, companyCode);
            WaitForLoad();

            _connection_string = WaitForElementIsVisible(By.Id(SFTP_CONNECTION_STRING));
            _connection_string.SetValue(PageBase.ControlType.TextBox, sftpConnectionString);
            WaitForLoad();

            _sftp_folder = WaitForElementIsVisible(By.Id(SFTP_FOLDER));
            _sftp_folder.SetValue(PageBase.ControlType.TextBox, sftpFolder);
            WaitForLoad();

            _is_active = WaitForElementExists(By.Id(IS_ACTIVE));
            _is_active.SetValue(PageBase.ControlType.CheckBox, isActive);
            WaitForLoad();

            _create_btn = WaitForElementIsVisible(By.Id(CREATE_BTN));
            _create_btn.Click();
            WaitPageLoading();
        }
        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
        }
        //---------Modal Create New Bank ----------
        public void SetName(string name)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.Clear();
            _name.SetValue(PageBase.ControlType.TextBox, name);
            WaitForLoad();
        }
        public void SetCronExpression(string cronExpression)
        {
            cron_expression = WaitForElementIsVisible(By.XPath(CRON_EXPRESSION));
            cron_expression.SetValue(PageBase.ControlType.TextBox, cronExpression);
            WaitForLoad();
        }
        public int GetNumberItemsDisplayed()
        {
            PageSize("100");
            var items = _webDriver.FindElements(By.XPath(ROWS_NUMBER));
            return items.Count();
        }

        public List<string> GetSelectedCustomersToFilter()
        {
            var customersSelected = new List<string>();
            _customer = WaitForElementIsVisible(By.Id(CUSTOMERS));
            _customer.Click();
            WaitForLoad();
            var selectElement = _webDriver.FindElements(By.XPath(CUSTOMERS_SELECTED));
            for (var i = 0; i < selectElement.Count; i++)
            {
                if (selectElement[i].Selected)
                {
                    var element = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[10]/ul/li[{0}]/label", i + 1)));
                    customersSelected.Add(element.Text);
                }
            }
            _customer.Click();
            return customersSelected;
        }
        public List<string> GetConverterTypesSelectedToFilter()
        {
            var converterTypesSelected = new List<string>();
            _converter_type = WaitForElementIsVisible(By.Id(CONVERTER_TYPES));
            _converter_type.Click();
            WaitForLoad();
            var selectElement = _webDriver.FindElements(By.XPath(CONVERTER_TYPES_SELECTED));
            for (var i = 0; i < selectElement.Count; i++)
            {
                if (selectElement[i].Selected)
                {
                    var element = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[11]/ul/li[{0}]/label/input/../span", i + 1)));
                    converterTypesSelected.Add(element.Text);
                }
            }

            _converter_type.Click();
            return converterTypesSelected;
        }

        public string GetShowFilterSelected()
        {

            var itemSelected = _webDriver.FindElements(By.XPath(SHOW_SELECTED));
            var showName = string.Empty;
            foreach (var element in itemSelected)
            {
                if (element.Selected)
                {
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
                        default:
                            showName = "";
                            break;
                    }
                }
            }

            return showName;
        }

        public void ClickCreate()
        {
            _create_btn = WaitForElementIsVisible(By.Id(CREATE_BTN));
            _create_btn.Click();
            WaitLoading();
        }


        public bool VerifyCreateValidators()
        {
            if (isElementVisible(By.XPath(NAME_VALIDATOR)) &&
                isElementVisible(By.XPath(SFTP_CONNECTION_STRING_VALIDATOR)) &&
                isElementVisible(By.XPath(SFTP_FOLDER_VALIDATOR)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void DeleteFirstBank()
        {
            var deleteBtn = WaitForElementExists(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[9]/a[2]"));
            deleteBtn.Click();
            WaitForLoad();

            var confirmDeleteBtn = WaitForElementExists(By.XPath("//*[@id=\"dataConfirmOK\"]"));
            confirmDeleteBtn.Click();
        }

        public string GetFirstLineName()
        {
            _firstLineName = WaitForElementExists(By.XPath(FIRST_LINE_NAME));
            WaitForLoad();
            return _firstLineName.Text.Trim();
        }

        public string GetFirstLineCustomer()
        {
            _firstLineCustomer = WaitForElementExists(By.XPath(FIRST_LINE_CUSTOMER));
            WaitForLoad();
            return _firstLineCustomer.Text.Trim();
        }

        public string GetFirstLineCompanyCode()
        {
            _firstLineCompanyCode = WaitForElementExists(By.XPath(FIRST_LINE_COMPANY_CODE));
            WaitForLoad();
            return _firstLineCompanyCode.Text.Trim();
        }

        public string GetFirstLineConverterType()
        {
            _firstLineConverter = WaitForElementExists(By.XPath(FIRST_LINE_CONVERTER));
            WaitForLoad();
            return _firstLineConverter.Text.Trim();
        }

        public void ClickFirstLine()
        {
            _firstJobRow = WaitForElementExists(By.XPath(FIRST_JOB_ROW));
            _firstJobRow.Click();
            WaitForLoad();
        }


        public int GetTotalResultRowsPage()
        {
            var rowsResults = _webDriver.FindElements(By.XPath(RESULTS_ROWS));
            WaitForLoad();
            return rowsResults.Count;
        }
        public List<string> GetAllPageResultConverterType()
        {
            var converterTypeList = new List<string>();

            var elements = _webDriver.FindElements(By.XPath(COL_CONVERTER_TYPE));
            if (elements != null)
            {
                foreach (var element in elements)
                {
                    converterTypeList.Add(element.Text.Trim());
                }
                return converterTypeList;
            }
            return new List<string>();
        }
        public List<string> GetAllPageResultCustomers()
        {
            var customersList = new List<string>();

            var elements = _webDriver.FindElements(By.XPath(COL_CUSTOMERS));
            if (elements != null)
            {
                foreach (var element in elements)
                {
                    customersList.Add(element.Text.Trim());
                }
                return customersList;
            }
            return new List<string>();
        }
        // to check All rows is active pass 'true'  -- to check all rows IsInactive pass 'false'
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

        public List<string> GetAllNamesPageResult()
        {
            var names = new List<string>();

            var elements = _webDriver.FindElements(By.XPath(NAMES_COL));
            if (elements == null)
            {
                return new List<string>();
            }
            foreach (var element in elements)
            {
                names.Add(element.Text.Trim());
            }
            return names;

        }
        public void EditorCronExpression()
        {
            _editor = WaitForElementIsVisible(By.Id(EDITOR));
            _editor.Click();
        }
        public bool CheckTabsExistance()
        {
            return isElementVisible(By.XPath(SECOND_TAB)) && isElementVisible(By.XPath(MINUTE_TAB))
                && isElementVisible(By.XPath(HOUR_TAB)) && isElementVisible(By.XPath(DAY_TAB))
                && isElementVisible(By.XPath(MONTH_TAB)) && isElementVisible(By.XPath(YEAR_TAB));
        }
        public string GetFirstFrequency()
        {
            _firstRowFrequency = WaitForElementExists(By.XPath(FIRST_ROW_FREQUENCY));
            return _firstRowFrequency.Text;
        }
        public BankCreateModalPage EditBank()
        {
            _firstJobRow = WaitForElementIsVisible(By.XPath(FIRST_JOB_ROW));
            _firstJobRow.Click();
            WaitForLoad();
            return new BankCreateModalPage(_webDriver, _testContext);
        }
        public int NombreRow()
        {
            var row = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr"));
            return row.Count;

        }
    }
}
