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
    public class JobNotifications : PageBase
    {

        private const string LINK_ADD_CONVERT_JOB_NOTIFICATION = "//*[@id=\"nav-tab\"]/li[8]/div/div/div[2]/a";
        private const string ADD_FILE_FLOW_JOB_NOTIFICATION = "//*[@id=\"nav-tab\"]/li[8]/div/div/div[3]/a";
        private const string Plus_BTN_FILEFLOWTYPE = "//*[@id=\"nav-tab\"]/li[8]/div/button";
        private const string RESET_FILTER_DEV = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string SEARCH_EMAIL = "JobNotificationFilters_SearchPattern";
        private const string DELETE = "dataConfirmOK";
        private const string CUSTOMER_FILTER = "JobNotificationFiltersSelectedCustomers_ms";
        private const string SEARCH_PATTERN = "JobNotificationFilters_SearchPattern";

        private const string CUSTOMERS_FILTER = "JobNotificationFiltersSelectedCustomers_ms";
        private const string UNSELECT_ALL_CUSTOMERS = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SELECT_ALL_CUSTOMERS = "/html/body/div[10]/div/ul/li[1]/a/span[2]";
        private const string CUSTOMERS_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string CUSTOMERS_XPATH = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[4]";
        private const string JOB_NOTIFICATIONS_NAMES = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[2]/text()";
        private const string PAGE_TWO = "//*[@id=\"list-item-with-action\"]/nav/ul/li[3]/a";
        private const string NOTIFICATION_TYPES_FILTER = "JobNotificationFiltersSelectedNotificationTypes_ms";
        private const string JOB_NOTIFICATIONS_TYPE = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[3]/text()";

        private const string NOTIFICATIONTYPE_FILTER = "JobNotificationFiltersSelectedNotificationTypes_ms";
        private const string UNSELECT_ALL_NOTIFICATIONTYPE = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string SELECT_ALL_NOTIFICATIONTYPE = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string NOTIFICATIONTYPE_SEARCH = "/html/body/div[11]/div/div/label/input";
        private const string NOTIFICATIONTYPE_XPATH = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[3]";

        private const string ADD_CUSTOMER_JOB_NOTIFICATION = "//*[@id=\"nav-tab\"]/li[8]/div/div/div[1]/a";

        private const string NEW_SAGEJOB = "//*[@id=\"nav-tab\"]/li[8]/div/div/div[4]/a";


        private const string LINK_ADD_CUSTOMER_JOB_NOTIFICATION = "//*[@id=\"nav-tab\"]/li[8]/div/div/div[1]/a";
        private const string NEW_SECIFICJOB = "//*[@id=\"nav-tab\"]/li[8]/div/div/div[5]/a";

        private const string CREATE_NEW_SAGE = "last";
        private const string JOBS_NOTIFICATIONS_ROWS = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr[*]/td[2]";
        private const string SECOND_PAGE = "//*[@id=\"list-item-with-action\"]/nav/ul/li[4]";
        private const string NEXT_PAGE = "//*[@id=\"list-item-with-action\"]/nav/ul/li[3]/a";
        private const string JOB_NOTIFICATION_TABLE = "//*[@id='list-item-with-action']/div[*]/div[*]/div[*]/div[2]/table/tbody/tr/td[2]";
        private const string JOB_NOTIFICATION_TABLE_TYPE = "//*[@id=\'list-item-with-action\']/div[*]/div[*]/div/div[2]/table/tbody/tr/td[3]";
        private const string JOB_NOTIFICATION_TABLE_LINE = "//*[@id=\"list-item-with-action\"]/div[*]/div[*]/div/div[2]/table/tbody/tr";
        private const string JOB_NOTIFICATIONS_CUSTOMER = "//*[@id=\"list-item-with-action\"]/div[*]/div[*]/div[*]/div[2]/table/tbody/tr[*]/td[4]";

        [FindsBy(How = How.XPath, Using = JOB_NOTIFICATION_TABLE_LINE)]
        private IWebElement _jobNotificationTableLine;

        [FindsBy(How = How.XPath, Using = JOB_NOTIFICATION_TABLE_TYPE)]
        private IWebElement _jobNotificationTableType;

        [FindsBy(How = How.XPath, Using = JOB_NOTIFICATION_TABLE)]
        private IWebElement _jobNotificationTable;

        [FindsBy(How = How.XPath, Using = LINK_ADD_CUSTOMER_JOB_NOTIFICATION)]
        private IWebElement _linkaddcustomerjobnotification;

        [FindsBy(How = How.XPath, Using = SECOND_PAGE)]
        private IWebElement _secondPage ;

        [FindsBy(How = How.XPath, Using = NEXT_PAGE)]
        private IWebElement _nextPage;

        [FindsBy(How = How.XPath, Using = LINK_ADD_CONVERT_JOB_NOTIFICATION)]
        private IWebElement _linkaddconvertjobnotification;

        [FindsBy(How = How.XPath, Using = ADD_FILE_FLOW_JOB_NOTIFICATION)]
        private IWebElement _addFileFlowJobNotification;

        [FindsBy(How = How.XPath, Using = Plus_BTN_FILEFLOWTYPE)]
        private IWebElement _plusbtnfileflowtype;

        [FindsBy(How = How.XPath, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.Id, Using = SEARCH_EMAIL)]
        private IWebElement _searchemail;

        [FindsBy(How = How.XPath, Using = DELETE)]
        private IWebElement _delete;

        [FindsBy(How = How.XPath, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_PATTERN)]
        private IWebElement _searchPattern;

        [FindsBy(How = How.Id, Using = CUSTOMERS_FILTER)]
        private IWebElement _customersFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_CUSTOMERS)]
        private IWebElement _unselectAllCustomers;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_CUSTOMERS)]
        private IWebElement _selectAllCustomers;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_SEARCH)]
        private IWebElement _searchCustomers;

        [FindsBy(How = How.Id, Using = NOTIFICATIONTYPE_FILTER)]
        private IWebElement _notificationTypeFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_NOTIFICATIONTYPE)]
        private IWebElement _unselectAllnotificationType;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_NOTIFICATIONTYPE)]
        private IWebElement _selectAllnotificationType;

        [FindsBy(How = How.XPath, Using = NOTIFICATIONTYPE_SEARCH)]
        private IWebElement _searchnotificationType;

        [FindsBy(How = How.XPath, Using = ADD_CUSTOMER_JOB_NOTIFICATION)]
        private IWebElement _addCustomerJobNotification;
        [FindsBy(How = How.XPath, Using = NEW_SAGEJOB)]
        private IWebElement _newsagejob;
        [FindsBy(How = How.Id, Using = CREATE_NEW_SAGE)]
        private IWebElement _createnewsagejob;



        public JobNotifications(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public void ShowBtnPlus()
        {
            _plusbtnfileflowtype = WaitForElementExists(By.XPath(Plus_BTN_FILEFLOWTYPE));
            _plusbtnfileflowtype.Click();
            WaitForLoad();
        }
        public ConvertJobNotificationCreateModalPage AddNewConverterJobNotification()
        {
            ShowBtnPlus();
            _linkaddconvertjobnotification = WaitForElementIsVisible(By.XPath(LINK_ADD_CONVERT_JOB_NOTIFICATION));
            _linkaddconvertjobnotification.Click();
            return new ConvertJobNotificationCreateModalPage(_webDriver, _testContext);
        }
        public FileFlowJobNotificationCreateModalPage AddNewFileFlowJobNotification()
        {
            ShowBtnPlus();
            _addFileFlowJobNotification = WaitForElementIsVisible(By.XPath(ADD_FILE_FLOW_JOB_NOTIFICATION));
            _addFileFlowJobNotification.Click();
            return new FileFlowJobNotificationCreateModalPage(_webDriver, _testContext);
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
        public enum FilterType
        {
            Search,
            Customer,
            NotificationTypes
        }
        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchemail = WaitForElementIsVisible(By.Id(SEARCH_EMAIL));
                    _searchemail.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                case FilterType.Customer:
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
                case FilterType.NotificationTypes:
                    ComboBoxSelectById(new ComboBoxOptions(NOTIFICATION_TYPES_FILTER, (string)value));
                    break;


                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }
            WaitPageLoading();
            WaitForLoad();
        }






        public void DeleteFirstJob()
        {
            var firstJobDelete = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[5]/a"));
            firstJobDelete.Click();
            _delete = WaitForElementIsVisible(By.Id(DELETE));
            _delete.Click();
            WaitForLoad();
        }
        public int CountJobs()
        {
            var firstJobDelete = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div[*]/div[*]/div[*]/div[2]/table/tbody/tr"));
            var count = firstJobDelete.Count();
            WaitForLoad();
            return count;
        }
        public void FilterByCutomer(string customer)
        {
            ComboBoxSelectById(new ComboBoxOptions(CUSTOMER_FILTER, customer));
        }
        public void FilterByEmail(string customer)
        {
            _searchPattern = WaitForElementIsVisible(By.Id(SEARCH_PATTERN));
            _searchPattern.SetValue(ControlType.TextBox, customer);
            WaitPageLoading();
            WaitLoading();
        }
        public List<string> GetEmailList()
        {
            var emailElements = _webDriver.FindElements(By.XPath(JOB_NOTIFICATION_TABLE));
            var emailList = emailElements.Select(element => element.Text).Where(text => IsValidEmail(text)).ToList();
            WaitForLoad();
            return emailList;
                    }
        public List<string> GetJobNotifLine()
        {
            var emailElements = _webDriver.FindElements(By.XPath(JOB_NOTIFICATION_TABLE_LINE));
            var emailList = emailElements.Select(element => element.Text).ToList();
            WaitForLoad();
            return emailList;
        }
        public List<string> GetTypeList()
        {
            var typeElements = _webDriver.FindElements(By.XPath(JOB_NOTIFICATION_TABLE_TYPE));
            var emailList = typeElements.Select(element => element.Text).ToList();
            WaitForLoad();
            return emailList;
        }
        private bool IsValidEmail(string email)
        {
            return !string.IsNullOrEmpty(email) && email.Contains("@");
        }


        public bool FindRowEmail(string email)
        {

            var fileFlowRow = _webDriver.FindElements(By.XPath(JOBS_NOTIFICATIONS_ROWS));
            foreach (var item in fileFlowRow)
            {
                if (item.Text.Equals(email))
                {
                    return true;
                }
            }
            return false;

        }

        public List<string> GetJobNotificationsList()
        {
            var jobNotificationsListId = new List<string>();

            var jobNotificationsList = _webDriver.FindElements(By.XPath(JOB_NOTIFICATIONS_NAMES));

            foreach (var jobNotification in jobNotificationsList)
            {
                jobNotificationsListId.Add(jobNotification.Text);
            }

            return jobNotificationsListId;
        }

        public void GoToPageTwo()
        {

            _secondPage = WaitForElementExists(By.XPath(SECOND_PAGE));
            _secondPage.Click();
                  WaitPageLoading();
        }

        public void GoToNextPage()
        {

            _nextPage = WaitForElementExists(By.XPath(NEXT_PAGE));
            _nextPage.Click();
            WaitPageLoading();
        }

        public CreateCustomerJobNotificationModel AddNewCustomerJobNotification()
        {
            ShowBtnPlus();
            _linkaddcustomerjobnotification = WaitForElementIsVisible(By.XPath(LINK_ADD_CUSTOMER_JOB_NOTIFICATION));
            _linkaddcustomerjobnotification.Click();
            return new CreateCustomerJobNotificationModel(_webDriver, _testContext);
        }
        public CreateSageJobNotificationModel AddNewSageJobNotification()
        {
            ShowBtnPlus();
            _newsagejob = WaitForElementIsVisible(By.XPath(NEW_SAGEJOB));
            _newsagejob.Click();
            return new CreateSageJobNotificationModel(_webDriver, _testContext);

        }

        public CreateSpecificJobNotificationModel AddNewSpecificJobNotification()
        {
            ShowBtnPlus();
            var _addNewSpecificJobnotification = WaitForElementIsVisible(By.XPath(NEW_SECIFICJOB));
            _addNewSpecificJobnotification.Click();
            return new CreateSpecificJobNotificationModel(_webDriver, _testContext);
        }

        public bool VerifyAllJobsBelongToCustomer(string customer)
        {

            var customerRow = _webDriver.FindElements(By.XPath(JOB_NOTIFICATIONS_CUSTOMER));
            foreach (var item in customerRow)
            {
                if (!item.Text.Equals(customer))
                {
                    return false;
                }
            }
            return true;

        }
    }
}

