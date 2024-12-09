using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static Newrest.Winrest.FunctionalTests.PageObjects.Shared.PageBase;
using DocumentFormat.OpenXml.Bibliography;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public class FileFlowProviderPage : PageBase
    {
        private const string Plus_BTN_FILEFLOWPROVIDER = "//*[@id=\"nav-tab\"]/li[8]/div/button";
        private const string Link_BTN_FILEFLOWPROVIDER = "//*[@id=\"nav-tab\"]/li[8]/div/div/div/a";  
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string SEARCH_NAME = "NameFolderEntityFilters_SearchPattern";
        private const string LINK_ADD_CONVERT_JOB_NOTIFICATION = "//*[@id=\"nav-tab\"]/li[8]/div/div/div[2]/a";
        private const string SEARCH_BY_NAME_VALUE = "//*[@id=\"NameFolderEntityFilters_SearchPattern\"]";
        private const string FILEFLOWPROVIDERS = "//*[@id=\"nav-tab\"]/li[1]/a";
        private const string DELETE_BUTTON = "dataConfirmOK";

        //__________________________________Variables_______________________________________

        [FindsBy(How = How.XPath, Using = Plus_BTN_FILEFLOWPROVIDER)]
        private IWebElement _plusbtnfileflowprovider;

        [FindsBy(How = How.XPath, Using = Link_BTN_FILEFLOWPROVIDER)]
        private IWebElement _linkbtnfileflowprovider;

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.Id, Using = SEARCH_NAME)]
        private IWebElement _searchname;

        [FindsBy(How = How.XPath, Using = LINK_ADD_CONVERT_JOB_NOTIFICATION)]
        private IWebElement _linkaddconvertjobnotification;
        
        [FindsBy(How = How.XPath, Using = SEARCH_BY_NAME_VALUE)]
        private IWebElement _searchbynamevalue;
        [FindsBy(How = How.XPath, Using = FILEFLOWPROVIDERS)]
        private IWebElement _fileflowproviders;


        public FileFlowProviderPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void ShowBtnPlus()
        {
            _plusbtnfileflowprovider = WaitForElementIsVisible(By.XPath(Plus_BTN_FILEFLOWPROVIDER));
            _plusbtnfileflowprovider.Click();
        }
        public FileFlowProviderCreateModalPage CreateModelFileFlowProvider()
        {
            ShowBtnPlus();

            _linkbtnfileflowprovider = WaitForElementIsVisible(By.XPath(Link_BTN_FILEFLOWPROVIDER));
            _linkbtnfileflowprovider.Click();

            return new FileFlowProviderCreateModalPage(_webDriver, _testContext);
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
        public enum FilterType
        {
            Search
        }
        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchname = WaitForElementIsVisible(By.Id(SEARCH_NAME));
                    _searchname.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }
            WaitPageLoading();
        }
        public int NombreRowFileFlowProvider()
        {
            var fileflowproviderrow = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr"));
            return fileflowproviderrow.Count;
        }
        public FileFlowProviderCreateModalPage AddNewConverterJobNotification()
        {
            ShowBtnPlus();

            _linkaddconvertjobnotification = WaitForElementIsVisible(By.XPath(LINK_ADD_CONVERT_JOB_NOTIFICATION));
            _linkaddconvertjobnotification.Click();

            return new FileFlowProviderCreateModalPage(_webDriver, _testContext);
        }

        public bool FindRowFileFlowProviderName(string name)
        {
            var fileflowproviderrow = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[2]"));
            foreach (var item in fileflowproviderrow)
            {
                if (item.Text.Equals(name))
                {
                    return true;
                }
            }
            return false;
        }

        public bool FindRowFileFlowProviderFolder(string folder)
        {
            var fileflowproviderrow = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[3]"));
            foreach (var item in fileflowproviderrow)
            {
                if (item.Text.Equals(folder))
                {
                    return true;
                }
            }
            return false;
        }

        public FileFlowProviderCreateModalPage ClickFirstRow()
        {
            var fileflowproviderrow = _webDriver.FindElement(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[1]"));
            fileflowproviderrow.Click();
            return new FileFlowProviderCreateModalPage(_webDriver, _testContext);
        }
        public void DeleteFirstRow()
        {
            var fileflowproviderrowDelete = _webDriver.FindElement(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[4]/a"));
            fileflowproviderrowDelete.Click();
            var deleteButton = _webDriver.FindElement(By.Id(DELETE_BUTTON));
            deleteButton.Click();
        }
        public string GetFileflowProvidersTitle()
        {
            var pagetitle = WaitForElementIsVisible(By.XPath(FILEFLOWPROVIDERS));
            return pagetitle.Text;
        }
        public string GetInputSearchValue()
        {
             _searchbynamevalue = WaitForElementIsVisible(By.XPath(SEARCH_BY_NAME_VALUE));
            return _searchbynamevalue.Text;
        }
        
    }

}
