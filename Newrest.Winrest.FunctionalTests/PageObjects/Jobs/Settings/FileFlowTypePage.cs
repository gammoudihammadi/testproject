using Limilabs.Client.IMAP;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
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
    public class FileFlowTypePage : PageBase
    {
        private const string Plus_BTN_FILEFLOWTYPE = "//*[@id=\"nav-tab\"]/li[8]/div/button";
        private const string Link_BTN_FILEFLOWTYPE = "//*[@id=\"nav-tab\"]/li[8]/div/div/div/a";
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string SEARCH_NAME = "NameFolderEntityFilters_SearchPattern";
        private const string LINK_ADD_CONVERT_JOB_NOTIFICATION = "//*[@id=\"nav-tab\"]/li[8]/div/div/div[2]/a";

        private const string FILE_FLOW_TYPES_NAMES = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[2]";
        //__________________________________Variables_______________________________________


        [FindsBy(How = How.XPath, Using = Plus_BTN_FILEFLOWTYPE)]
        private IWebElement _plusbtnfileflowtype;

        [FindsBy(How = How.XPath, Using = Link_BTN_FILEFLOWTYPE)]
        private IWebElement _linkbtnfileflowtype;

        [FindsBy(How = How.XPath, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.Id, Using = SEARCH_NAME)]
        private IWebElement _searchname;

        [FindsBy(How = How.XPath, Using = LINK_ADD_CONVERT_JOB_NOTIFICATION)]
        private IWebElement _linkaddconvertjobnotification;

        public FileFlowTypePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void ShowBtnPlus()
        {
            _plusbtnfileflowtype = WaitForElementIsVisible(By.XPath(Plus_BTN_FILEFLOWTYPE));
            _plusbtnfileflowtype.Click();           
        }
        public FileFlowTypeCreateModalPage CreateModelFileFlowType()
        {
            ShowBtnPlus();

            _plusbtnfileflowtype = WaitForElementIsVisible(By.XPath(Link_BTN_FILEFLOWTYPE));
            _linkbtnfileflowtype.Click();

            return new FileFlowTypeCreateModalPage(_webDriver, _testContext);
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
            Thread.Sleep(1500);
        }
        public int NombreRowFileFlowType()
        {

            var fileflowtyperow = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr"));
            return fileflowtyperow.Count;

        }

        public FileFlowTypeCreateModalPage AddNewConverterJobNotification()
        {
            ShowBtnPlus();

            _linkaddconvertjobnotification = WaitForElementIsVisible(By.XPath(LINK_ADD_CONVERT_JOB_NOTIFICATION));
            _linkaddconvertjobnotification.Click();

            return new FileFlowTypeCreateModalPage(_webDriver, _testContext);
        }

        public bool FindRowFileFlowTypeName(string name)
        {

            var fileflowtyperow = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[2]"));
            foreach (var item in fileflowtyperow)
            {
                if (item.Text.Equals(name))
                {
                    return true;
                }
            }
            return false;

        }

        public bool FindRowFileFlowTypeFolder(string folder)
        {

            var fileflowtyperow = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[3]"));
            foreach (var item in fileflowtyperow)
            {
                if (item.Text.Equals(folder))
                {
                    return true;
                }
            }
            return false;

        }

        
        public FileFlowProvidersEditModal ClickFirstRow()
        {

            var fileflowtyperow = _webDriver.FindElement(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[1]"));
            fileflowtyperow.Click();
            return  new FileFlowProvidersEditModal(_webDriver, _testContext);

        }
        public void DeleteFirstRow()
        {

            var fileflowtyperowDelete = _webDriver.FindElement(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[4]/a"));
            fileflowtyperowDelete.Click();
            var deleteButton = _webDriver.FindElement(By.XPath("//*[@id=\"dataConfirmOK\"]"));
            deleteButton.Click();
        }

        public List<string> GetFileFlowTypesList()
        {
            var fileFlowTypesListId = new List<string>();

            var fileFlowTypesList = _webDriver.FindElements(By.XPath(FILE_FLOW_TYPES_NAMES));

            foreach (var fileFlowType in fileFlowTypesList)
            {
                fileFlowTypesListId.Add(fileFlowType.Text.Remove(0, 3));
            }

            return fileFlowTypesListId;
        }
    }
}
