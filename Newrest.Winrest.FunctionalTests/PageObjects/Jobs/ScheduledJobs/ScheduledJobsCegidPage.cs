using DocumentFormat.OpenXml.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs
{
    public class ScheduledJobsCegidPage : PageBase
    {
        public ScheduledJobsCegidPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }
        //___________________________________ Constantes _____________________________________________

        private const string PLUS_BTN = "//*[@id=\"nav-tab\"]/li[5]/div/button";
        private const string NEW_CEGID_SCHED_JOB = "//*[@id=\"nav-tab\"]/li[5]/div/div/div/a";
        private const string SEARCH_BY_Name = "SearchPattern";
        private const string SHOW_ACTIVE_FILTER = "ShowOnlyActive";
        private const string RESET_FILTER = "ResetFilter";
        private const string NAMES_COLUMNS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[2]";
        private const string LIST_ROWS = "//*[@id='tableListMenu']/tbody/tr[*]";
        private const string DELETE_FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr/td[7]/a[2]";
        private const string DELETE_CONFIRM = "dataConfirmOK";
        private const string CEGIDNAME = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[2]";
        private const string CEGID_JOB_TYPE_FILTRE = "SageScheduledJobFiltersSelectedSageJobTypes_ms";
        private const string CONVERTER_FILTER = "FileFlowScheduledJobFiltersSelectedConverters_ms";
        private const string UNSELECT_ALL_CEGID_JOB_TYPE = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SELECT_ALL_CEGID_JOB_TYPE = "/html/body/div[10]/div/ul/li[1]/a/span[2]";
        private const string CEGIDJOBTYPE_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string CEGIDJOBTYPE_XPATH = "//*[starts-with(@id,\"tableListMenu\")]/tbody/tr/td[6]";
        private const string SHOW_INACTIVE_ONLY = "ShowOnlyInactive";
        private const string SHOW_ALL = "ShowAll";
        private const string MODAL_CONFIRM_DELETE = "//*[@id=\"dataConfirmModal\"]/div/div";
        private const string PROVIDERS_NAMES = "//*[@id=\"tableListMenu\"]/thead/tr/th[2]";
        private const string FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr[1]";
        private const string POPUP = "modal-1";
        private const string POPUP_CHECK = "//*[@class=\"modal-dialog\"]";
        private const string FIRST_ROW_FREQUENCY = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[3]";

        //tableau
        private const string RESULTS_ROWS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]";



        //_____________________________________ Variables _____________________________________________
        // General
        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusBtn;

        [FindsBy(How = How.XPath, Using = NEW_CEGID_SCHED_JOB)]
        private IWebElement _newCegidSchedJob;

        [FindsBy(How = How.Id, Using = CEGIDNAME)]
        private IWebElement _cegid;

        [FindsBy(How = How.XPath, Using = DELETE_FIRST_ROW)]
        private IWebElement _deleteFirstRow;

        [FindsBy(How = How.Id, Using = DELETE_CONFIRM)]
        private IWebElement _deleteConfirm;

        [FindsBy(How = How.XPath, Using = FIRST_ROW)]
        private IWebElement _firstRow;
        [FindsBy(How = How.Id, Using = POPUP)]
        private IWebElement _popup;

        // Filters

        [FindsBy(How = How.Id, Using = SEARCH_BY_Name)]
        private IWebElement _searchNameFilter;

        [FindsBy(How = How.Id, Using = SHOW_ACTIVE_FILTER)]
        private IWebElement _showActiveOnlyFilter;

        [FindsBy(How = How.Id, Using = SHOW_INACTIVE_ONLY)]
        private IWebElement _showInactiveOnlyFilter;

        [FindsBy(How = How.Id, Using = SHOW_ALL)]
        private IWebElement _showALLFilter;

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = CEGID_JOB_TYPE_FILTRE)]
        private IWebElement _cegidJobTypeFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_CEGID_JOB_TYPE)]
        private IWebElement _unselectAllCegidJobType;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_CEGID_JOB_TYPE)]
        private IWebElement _selectAllCegidJobType;

        [FindsBy(How = How.XPath, Using = CEGIDJOBTYPE_SEARCH)]
        private IWebElement _searchCegidJobType;

        [FindsBy(How = How.XPath, Using = FIRST_ROW_FREQUENCY)]
        private IWebElement _firstRowFrequency;

        public enum FilterType
        {
            SearchByName,
            ShowAll,
            ShowActiveOnly,
            ShowInactiveOnly,

            GegidJobTypes,
            GegidJobTypesUncheck,
            GegidJobTypesCheckAll,

        }

        public void Filter(FilterType FilterType, object value)
        {
            switch (FilterType)
            {
                case FilterType.SearchByName:
                    _searchNameFilter = WaitForElementIsVisible(By.Id(SEARCH_BY_Name));
                    _searchNameFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    WaitPageLoading();
                    break;
                case FilterType.ShowActiveOnly:
                    _showActiveOnlyFilter = WaitForElementIsVisible(By.Id(SHOW_ACTIVE_FILTER));
                    _showActiveOnlyFilter.SetValue(ControlType.CheckBox, value);
                    WaitPageLoading();
                    break;
                case FilterType.ShowAll:
                    _showALLFilter = WaitForElementExists(By.Id(SHOW_ALL));
                    _showALLFilter.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;
                case FilterType.ShowInactiveOnly:
                    _showInactiveOnlyFilter = WaitForElementExists(By.Id(SHOW_INACTIVE_ONLY));
                    _showInactiveOnlyFilter.SetValue(ControlType.RadioButton, value);
                    WaitPageLoading();
                    break;

                case FilterType.GegidJobTypesUncheck:
                    _cegidJobTypeFilter = WaitForElementIsVisible(By.Id(CEGID_JOB_TYPE_FILTRE));
                    _cegidJobTypeFilter.Click();

                    _unselectAllCegidJobType = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_CEGID_JOB_TYPE));
                    _unselectAllCegidJobType.Click();

                    _cegidJobTypeFilter.Click();
                    break;

                case FilterType.GegidJobTypesCheckAll:
                    _cegidJobTypeFilter = WaitForElementIsVisible(By.Id(CEGID_JOB_TYPE_FILTRE));
                    _cegidJobTypeFilter.Click();

                    var selectAllCustomers = WaitForElementIsVisible(By.XPath(SELECT_ALL_CEGID_JOB_TYPE));
                    selectAllCustomers.Click();

                    _cegidJobTypeFilter.Click();
                    break;

                case FilterType.GegidJobTypes:
                    _cegidJobTypeFilter = WaitForElementIsVisible(By.Id(CEGID_JOB_TYPE_FILTRE));
                    _cegidJobTypeFilter.Click();

                    _unselectAllCegidJobType = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_CEGID_JOB_TYPE));
                    _unselectAllCegidJobType.Click();

                    _searchCegidJobType = WaitForElementIsVisible(By.XPath(CEGIDJOBTYPE_SEARCH));
                    _searchCegidJobType.SetValue(ControlType.TextBox, value);
                    WaitForLoad();

                    var cegidJobTypeToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    cegidJobTypeToCheck.SetValue(ControlType.CheckBox, true);

                    _cegidJobTypeFilter.Click();
                    break;
            }
            WaitPageLoading();
            WaitForLoad();
        }

        public CegidScheduledJobModal NewCegidScheduledModal()
        {
            _plusBtn = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            _plusBtn.Click();
            WaitForLoad();
            _newCegidSchedJob = WaitForElementExists(By.XPath(NEW_CEGID_SCHED_JOB));
            _newCegidSchedJob.Click();
            WaitForLoad();
            return new CegidScheduledJobModal(_webDriver, _testContext);
        }

        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
            }
        }
        public List<string> GetAllNamesResultPaged()
        {
            var names = new List<string>();

            var elements = _webDriver.FindElements(By.XPath(NAMES_COLUMNS));
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

        public void CLickOnStart()
        {

            Random random = new Random();
            int randomNumber = random.Next(1, GetCountResultPaged());
            var element = _webDriver.FindElement(By.XPath($"//*[@id=\"tableListMenu\"]/tbody/tr[{randomNumber}]/td[7]/a[1]"));
            element.Click();
            WaitForElementIsVisible(By.XPath("//*[@id=\"dataConfirmOK\"]")).Click();


        }

        public bool CheckStatutIsVisible()
        {
            var element = WaitForElementExists(By.XPath("//*[@id=\"div-body\"]/div/div[1]/div[3]"));
            return element != null;
        }

        public int GetCountResultPaged()
        {
            var elements = _webDriver.FindElements(By.XPath(LIST_ROWS));
            return elements.Count;
        }
        public void DeleteFirstRow()
        {
            if (isElementVisible(By.XPath(DELETE_FIRST_ROW)))
            {
                _deleteFirstRow = WaitForElementIsVisible(By.XPath(DELETE_FIRST_ROW));
                _deleteFirstRow.Click();

                _deleteConfirm = WaitForElementIsVisible(By.Id(DELETE_CONFIRM));
                _deleteConfirm.Click();
            }
        }

        public void ClickDelete()
        {
            _deleteFirstRow = WaitForElementIsVisible(By.XPath(DELETE_FIRST_ROW));
            _deleteFirstRow.Click();


        }
        public string GetFirstCegidName()
        {
            var _cegid = WaitForElementIsVisible(By.XPath(CEGIDNAME));
            return _cegid.Text.Trim();
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

            var elements = _webDriver.FindElements(By.XPath(NAMES_COLUMNS));
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
        public void ClickOnFirstRow()
        {
            _firstRow = WaitForElementIsVisible(By.XPath(FIRST_ROW));
            _firstRow.Click();
            WaitForLoad();
        }
        public CegidScheduledJobModal EditFirstRow()
        {
            _firstRow = WaitForElementIsVisible(By.XPath(FIRST_ROW));
            _firstRow.Click();
            WaitForLoad();
            return new CegidScheduledJobModal(_webDriver, _testContext);

        }
        public bool CheckOpenedPopup()
        {
            WaitForLoad();
            return isElementExists(By.Id(POPUP));
        }
        public CegidScheduledJobModal ClickFirstLine()
        {
            _firstRow = WaitForElementExists(By.XPath(FIRST_ROW));
            _firstRow.Click();
            WaitForLoad();
            return new CegidScheduledJobModal(_webDriver, _testContext);
        }
        public bool GegidJobTypesFiltred(string cegidType)
        {
            PageSize("100");
            var gegidJobTypes = _webDriver.FindElements(By.XPath(CEGIDJOBTYPE_XPATH));

            List<string> filteredCustomers = new List<string>();

            foreach (var elm in gegidJobTypes)
            {
                filteredCustomers.Add(elm.Text.Trim());
            }

            if (!filteredCustomers.Contains(cegidType))
            {
                return false;
            }
            return true;
        }

        public bool CheckStatus()
        {
            var allrowactive = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]"));

            foreach (var elm in allrowactive)
            {
                if (elm.GetAttribute("class").Contains("IsInactive"))
                {
                    return false;
                }
            }


            return true;
        }
        public int GetTotalResultRowsPage()
        {
            var rowsResults = _webDriver.FindElements(By.XPath(RESULTS_ROWS));
            WaitForLoad();
            return rowsResults.Count;
        }
        public string GetFirstLineName()
        {
            var firstLineName = WaitForElementExists(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[2]"));
            WaitForLoad();
            return firstLineName.Text;
        }
        public string GetFirstFrequency()
        {
            _firstRowFrequency = WaitForElementExists(By.XPath(FIRST_ROW_FREQUENCY));
            return _firstRowFrequency.Text;


        }

        public bool GoToNextPage(int c)
        {
            if (isElementExists(By.XPath($"//*[@id=\"list-item-with-action\"]/nav/ul/li[{c}]/a")))
            {
                WaitForElementExists(By.XPath($"//*[@id=\"list-item-with-action\"]/nav/ul/li[{c}]/a")).Click();
                WaitPageLoading();
                return false;
            }
           else
           {
                return true;
           }
           
        }

        public int GetTotalResultsAllPages()
        {
            var resultnumber = 0;

            for (int i = 3; i <= 100; i++)
            {
                resultnumber += GetTotalResultRowsPage();
                var result = GoToNextPage(i);
                if (result)
                {
                    break;
                }
            }

            return resultnumber;
        }

    }
}

