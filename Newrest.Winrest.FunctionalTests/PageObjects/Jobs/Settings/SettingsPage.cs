using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Internal;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public class SettingsPage : PageBase
    {


        //__________________________________Constantes_______________________________________


        private const string FILE_FLOW_TYPE = "//*[@id=\"nav-tab\"]/li[2]/a";
        private const string Plus_BTN_FILEFLOWTYPE = "//*[@id=\"nav-tab\"]/li[8]/div/button";
        private const string FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr[1]";
        private const string NAME_FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr/td[2]";
        private const string FOLDER_FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr/td[3]";
        //Filter
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string JOB_NOTIFICATION = "//*[@id=\"nav-tab\"]/li[6]/a";
        private const string CEGID_PANEL = "//*[@id=\"nav-tab\"]/li[7]/a";
        private const string FILE_FLOW_PROVIDER = "//*[@id=\"nav-tab\"]/li[1]/a";
        private const string CONVERTERS = "//*[@id=\"nav-tab\"]/li[3]/a";

        private const string Plus_BTN_FILEFLOWPROVIDERS = "//*[@id=\"nav-tab\"]/li[8]/div/button";
        private const string Link_BTN_FILEFLOWPROVIDERS = "//*[@id=\"nav-tab\"]/li[8]/div/div/div/a";

        private const string DELETE_ICONS = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[contains(normalize-space(text()), '{0}')]/../td[4]/a/span";
        private const string DELETE_BUTTONS = "dataConfirmOK";

        private const string LIST_ELEMENTS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[2]";

        private const string SEARCH_FILTER = "NameFolderEntityFilters_SearchPattern";

        private const string FILE_FLOW_PROVIDERS_NAMES = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[2]";

        private const string PAGE_TWO = "//*[@id=\"list-item-with-action\"]/nav/ul/li[4]";

        private const string GUEST_MAPPINGS = "//*[@id=\"nav-tab\"]/li[4]/a";
                //__________________________________Variables_______________________________________

        [FindsBy(How = How.XPath, Using = FILE_FLOW_TYPE)]
        private IWebElement _fileflowtype;

        [FindsBy(How = How.XPath, Using = Plus_BTN_FILEFLOWTYPE)]
        private IWebElement _plusbtnfileflowtype;

        [FindsBy(How = How.XPath, Using = Plus_BTN_FILEFLOWPROVIDERS)]
        private IWebElement _plusbtnfileflowproviders;

        [FindsBy(How = How.XPath, Using = Link_BTN_FILEFLOWPROVIDERS)]
        private IWebElement _linkbtnfileflowtype;


        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.XPath, Using = JOB_NOTIFICATION)]
        private IWebElement _jobnotification;

        [FindsBy(How = How.XPath, Using = CONVERTERS)]
        private IWebElement _converters;

        [FindsBy(How = How.XPath, Using = FILE_FLOW_PROVIDER)]
        private IWebElement _fileflowprovider;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.XPath, Using = CEGID_PANEL)]
        private IWebElement _cegidpanel;

        [FindsBy(How = How.XPath, Using = GUEST_MAPPINGS)]
        private IWebElement _guestMappings;
        public SettingsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
     
        public FileFlowTypePage GoToFILEFLOWPAGE()
        {
            WaitForLoad();
            _fileflowtype = WaitForElementIsVisible(By.XPath(FILE_FLOW_TYPE));
            _fileflowtype.Click();
            WaitForLoad();

            return new FileFlowTypePage(_webDriver, _testContext);
        }
        public JobNotifications GoTo_JOB_NOTIFICATION()
        {
            WaitForLoad();
            _jobnotification = WaitForElementIsVisible(By.XPath(JOB_NOTIFICATION));
            _jobnotification.Click();
            WaitForLoad();
            return new JobNotifications(_webDriver, _testContext);
        }

        public void ResetFilter()
        {
            _resetFilterDev = WaitForElementIsVisible(By.Id(RESET_FILTER_DEV));
            _resetFilterDev.Click();
            WaitForLoad();
        }

        public void ShowBtnPlus()
        {
            _plusbtnfileflowproviders = WaitForElementIsVisible(By.XPath(Plus_BTN_FILEFLOWPROVIDERS));
            _plusbtnfileflowproviders.Click();
        }

        public SettingsCreateModalPage CreateModelFileFlowProviders()
        {
            ShowBtnPlus();

            _plusbtnfileflowproviders = WaitForElementIsVisible(By.XPath(Link_BTN_FILEFLOWPROVIDERS));
            _linkbtnfileflowtype.Click();

            return new SettingsCreateModalPage(_webDriver, _testContext);
        }

        public void DeleteElementFileFlowProviders(string filteFlowProvider)
        {
            var DeleteButtons = WaitForElementIsVisible(By.XPath(String.Format(DELETE_ICONS, filteFlowProvider)));
            DeleteButtons.Click();
            var deleteButton = _webDriver.FindElement(By.Id(DELETE_BUTTONS));
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

        public FileFlowProvidersEditModal SelectFirstRow()
        {
            WaitForLoad();
            var firstRow = WaitForElementIsVisible(By.XPath(FIRST_ROW));
            firstRow.Click();
            WaitPageLoading();
            return new FileFlowProvidersEditModal(_webDriver, _testContext);
        }
        public string GetNameFirstRow()
        {
            var nameRow = WaitForElementIsVisible(By.XPath(NAME_FIRST_ROW));
            return nameRow.Text;
        }
        public string GetFolderFirstRow()
        {
            var folderRow = WaitForElementIsVisible(By.XPath(FOLDER_FIRST_ROW));
            return folderRow.Text;
        }

        public FileFlowProviderPage GoToFILEFLOWPROVIDERPAGE()
        {
            WaitForLoad();
            _fileflowprovider = WaitForElementIsVisible(By.XPath(FILE_FLOW_PROVIDER));
            _fileflowprovider.Click();
            WaitForLoad();

            return new FileFlowProviderPage(_webDriver, _testContext);
        }
        public SettingsTabConvertersPage  GoTo_CONVERTERS()
        {
            WaitForLoad();
            _converters = WaitForElementIsVisible(By.XPath(CONVERTERS));
            _converters.Click();
            WaitForLoad();
            return new SettingsTabConvertersPage(_webDriver, _testContext);
        }

        public enum FilterType
        {
            Search,
        }
        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    WaitLoading();
                    break;
                default:
                    break;
            }
            WaitPageLoading();
            WaitForLoad();
        }

        public List<string> GetFileFlowProvidersList()
        {
            var fileFlowProvidersListId = new List<string>();

            var fileFlowProvidersList = _webDriver.FindElements(By.XPath(FILE_FLOW_PROVIDERS_NAMES));

            foreach (var fileFlowProviders in fileFlowProvidersList)
            {
                fileFlowProvidersListId.Add(fileFlowProviders.Text);
            }

            return fileFlowProvidersListId;
        }

        public void GoToPageTwo()
        {
            var _pageTwo = WaitForElementExists(By.XPath(PAGE_TWO));
            _pageTwo.Click();
            WaitForLoad();
        }

        public GuestMappingsPage GoTo_GUEST_MAPPINGS()
        {
            WaitForLoad();
            _converters = WaitForElementIsVisible(By.XPath(GUEST_MAPPINGS));
            _converters.Click();
            WaitForLoad();
            return new GuestMappingsPage(_webDriver, _testContext);
        }   
    }
}
