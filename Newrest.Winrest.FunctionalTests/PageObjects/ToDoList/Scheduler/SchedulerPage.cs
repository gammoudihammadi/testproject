using DocumentFormat.OpenXml.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Schedule;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Tasks;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Scheduler
{
    public class SchedulerPage : PageBase
    {
        public SchedulerPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        //---------------------------------------CONSTANTES---------------------------------------------------------------

        public const string PLUS_BTN = "/html/body/div[2]/div/div[2]/div/div[1]/div/div/button";
        public const string CREATE_NEW_SCHEDULER = "/html/body/div[2]/div/div[2]/div/div[1]/div/div/div/a[1]";
        public const string DUPLICATE_SCHEDULER = "/html/body/div[2]/div/div[2]/div/div[1]/div/div/div/a[2]";
        public const string TASK_TAB = "/html/body/div[1]/div[1]/div[4]/ul/li[10]/div/div/ul/li[1]/a";
        public const string SCHEDULER_TAB = "/html/body/div[1]/div[1]/div[4]/ul/li[10]/div/div/ul/li[2]/a";
        public const string NUMBER_SCHEDULER_IN_HEADER = "/html/body/div[2]/div/div[2]/div/div[1]/h1/span";
        public const string BACK_TO_LIST = "/html/body/div[2]/a/span[1]";
        public const string DELETE_ICON = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[6]/a/span";
        public const string DELETE_CONFIRM = "//*/a[text()='Delete']";
        public const string FIRST_SCHEDULER_NAME = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr/td[2]";
        public const string FIRST_TASK_NAME = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr/td[3]";
        public const string FIRST_SITE = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr/td[4]";
        public const string FIRST_STATUS = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr/td[5]";
        private const string DETAIL_BUTTON = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[1]";
        private const string SCHEDULER_NAMES = "//*[@id='list-item-with-action']/table/tbody/tr[*]/td[2]";


        //---------------------------------------FILTERS CONSTANTES-------------------------------------------------------
        public const string SEARCH_SCHEDULER_NAME = "SearchString";
        public const string SEARCH_TASK_NAME = "TaskSearchString";
        public const string TEST = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[4]/input[1]";
        public const string PRODUCTION = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[4]/input[2]";
        public const string OBSOLETE = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[4]/input[3]";
        public const string SITES = "FilterSites_ms";
        public const string FROM = "date-picker-start";
        public const string TO = "date-picker-end";
        public const string RESET_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[1]/a";
        public const string LIST_SITE_FILTER = "/html/body/div[10]/ul/li[*]/label/input";
        public const string SITE_FIRST_ROW = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[4]";

        //---------------------------------------DUPLICATE CONSTANTES-------------------------------------------------------
        public const string SCHEDULER = "Schedulers";
        public const string TASK = "Tasks";
        public const string NEW_SCHEDULER_NAME = "SchedulerName";
        public const string DUPLICATE_SCHEDULER_BUTTON = "DuplicateTaskButton";

        //tableau
        private const string FIRST_ROW = "//*[@id=\"list-item-with-action\"]/table/tbody/tr";
        private const string SITE = "SelectedSites";
        private const string SITE_XPATH = "//*[starts-with(@id,\"list-item-with-action\")]/table/tbody/tr[*]/td[4]";

        // ____________________________________ Variables _______________________________________________
        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusBtn;

        [FindsBy(How = How.XPath, Using = CREATE_NEW_SCHEDULER)]
        private IWebElement _createBtn;

        [FindsBy(How = How.XPath, Using = TEST)]
        private IWebElement _test;

        [FindsBy(How = How.XPath, Using = PRODUCTION)]
        private IWebElement _production;

        [FindsBy(How = How.XPath, Using = OBSOLETE)]
        private IWebElement _obsolete;

        [FindsBy(How = How.Id, Using = FROM)]
        private IWebElement _from;

        [FindsBy(How = How.Id, Using = TO)]
        private IWebElement _to;

        [FindsBy(How = How.Id, Using = SEARCH_SCHEDULER_NAME)]
        private IWebElement _searchScheduler;

        [FindsBy(How = How.Id, Using = SEARCH_TASK_NAME)]
        private IWebElement _searchTask;

        [FindsBy(How = How.Id, Using = SCHEDULER)]
        private IWebElement _scheduler;

        [FindsBy(How = How.Id, Using = TASK)]
        private IWebElement _task;

        [FindsBy(How = How.Id, Using = NEW_SCHEDULER_NAME)]
        private IWebElement _newSchedulerName;

        [FindsBy(How = How.XPath, Using = DETAIL_BUTTON)]
        private IWebElement _detail_Button;

        //tableau
        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How =How.XPath, Using = SITE_FIRST_ROW)]
        private IWebElement _siteFirstRow;

        public enum FilterType
        {
            Test,
            Production,
            Obsolete,
            SearchSchedulerName,
            Sites,
            SearchTaskName,
            From,
            To
        }

        public void Filter(FilterType filterType, object value)

        {
            switch (filterType)
            {
                case FilterType.SearchSchedulerName:

                    _searchScheduler = WaitForElementIsVisible(By.Id(SEARCH_SCHEDULER_NAME));
                    _searchScheduler.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.SearchTaskName:

                    _searchTask = WaitForElementIsVisible(By.Id(SEARCH_TASK_NAME));
                    _searchTask.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.Test:
                    _test = WaitForElementIsVisible(By.XPath(TEST));
                    _test.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.Production:
                    _production = WaitForElementIsVisible(By.XPath(PRODUCTION));
                    _production.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.Obsolete:
                    _obsolete = WaitForElementIsVisible(By.XPath(OBSOLETE));
                    _obsolete.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.Sites:
                    ComboBoxSelectById(new ComboBoxOptions(SITES, (string)value));
                    break;

                case FilterType.From:
                    _from = WaitForElementIsVisible(By.Id(FROM));
                    _from.SetValue(ControlType.DateTime, value);
                    _from.SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;

                case FilterType.To:
                    _to = WaitForElementIsVisible(By.Id(TO));
                    _to.SetValue(ControlType.DateTime, value);
                    _to.SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;


                default:
                    break;

            }
            WaitPageLoading();
            Thread.Sleep(2000);
        }

        public object GetFilterValue(FilterType filterType)
        {
            switch (filterType)
            {
                case FilterType.SearchSchedulerName:
                    _searchScheduler = WaitForElementIsVisible(By.Id(SEARCH_SCHEDULER_NAME));
                    return _searchScheduler.GetAttribute("value");

                case FilterType.SearchTaskName:
                    _searchTask = WaitForElementIsVisible(By.Id(SEARCH_TASK_NAME));
                    return _searchTask.GetAttribute("value");

                case FilterType.Test:
                    _test = WaitForElementIsVisible(By.XPath(TEST));
                    return _test.Selected;

                case FilterType.Production:
                    _production = WaitForElementIsVisible(By.XPath(PRODUCTION));
                    return _production.Selected;

                case FilterType.Obsolete:
                    _obsolete = WaitForElementIsVisible(By.XPath(OBSOLETE));
                    return _obsolete.Selected;

                case FilterType.From:
                    _from = WaitForElementIsVisible(By.Id(FROM));
                    return _from.GetAttribute("value");

                case FilterType.To:
                    _to = WaitForElementIsVisible(By.Id(TO));
                    return _to.GetAttribute("value");
            }
            return null;
        }
        public int GetNumberSelectedSiteFilter()
        {
            var listSitesSelectedFilters = _webDriver.FindElements(By.XPath(LIST_SITE_FILTER));
            var numberSitesSelectedSite = listSitesSelectedFilters
               .Where(p => p.Selected == true).Count();

            return numberSitesSelectedSite;
        }
        public void ResetFilters()
        {
            var resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            resetFilter.Click();

            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //pas de date
            }
        }
        public void ShowPlusScheduler()
        {
            _plusBtn = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            _plusBtn.Click();
            WaitForLoad();
        }

        public SchedulerCreateModalPage CreateSchedulerModalPage()
        {
            ShowPlusScheduler();
            _createBtn = WaitForElementIsVisible(By.XPath(CREATE_NEW_SCHEDULER));
            _createBtn.Click();
            WaitForLoad();

            return new SchedulerCreateModalPage(_webDriver, _testContext);
        }
        public TasksPage GoToTaskTab() 
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var todoListMenu = WaitForElementIsVisible(By.Id("TabLounge"));
                todoListMenu.Click();

                var taskMenu = WaitForElementIsVisible(By.Id("TaskTabNav"));
                taskMenu.Click();
            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();

                var taskMenuLink = WaitForElementIsVisible(By.Id("TaskLinkDashBoard"));
                taskMenuLink.Click();
            }

            WaitForLoad();

            return new TasksPage(_webDriver, _testContext);
        }
        public SchedulerPage GoToSchedulerTab()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var todoListMenu = WaitForElementIsVisible(By.Id("TabLounge"));
                todoListMenu.Click();

                var schedulerMenu = WaitForElementIsVisible(By.Id("SchedulerTabNav"));
                schedulerMenu.Click();
            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();

                var schedulerMenuLink = WaitForElementIsVisible(By.Id("SchedulerLinkDashBoard"));
                schedulerMenuLink.Click();
            }

            WaitForLoad();

            return new SchedulerPage(_webDriver, _testContext);
        }
        public string GetNumberOfSchedulerInHeader()
        {
            var numberScheduler = WaitForElementIsVisible(By.XPath(NUMBER_SCHEDULER_IN_HEADER));
            return numberScheduler.Text;
        }
        public void BackToList()
        {
            var backToListBtn = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            backToListBtn.Click();
            WaitLoading();
        }
        public void DeleteScheduler()
        {
            var delete = WaitForElementIsVisible(By.XPath(DELETE_ICON));
            delete.Click();
            WaitForLoad();

            var deleteConfirm = WaitForElementIsVisible(By.XPath(DELETE_CONFIRM));
            deleteConfirm.Click();
            WaitForLoad();
        }
        public bool VerifySearchSchedulerFilter(string name)
        {
            var firstSchedulerName = WaitForElementIsVisible(By.XPath(FIRST_SCHEDULER_NAME));
            WaitPageLoading();
            if(firstSchedulerName.Text != name)
            {
                return false;
            }
            return true;
        }
        public bool VerifySchedulerNamesMatch(string name)
        {
            var schedulerNames = _webDriver.FindElements(By.XPath(SCHEDULER_NAMES));
            WaitPageLoading();

            foreach (var schedulerName in schedulerNames)
            {
                if (schedulerName.Text != name)
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifySearchTaskNameFilter(string name)
        {
            var firstTaskName = WaitForElementIsVisible(By.XPath(FIRST_TASK_NAME));
            if (firstTaskName.Text != name)
            {
                return false;
            }
            return true;
        }
        public bool VerifySiteFilter(string site)
        {
            var firstSite = WaitForElementIsVisible(By.XPath(FIRST_SITE));
            if (firstSite.Text != site)
            {
                return false;
            }
            return true;
        }
        public bool VerifyStatusFilter(string status)
        {
            var firstStatus = WaitForElementIsVisible(By.XPath(FIRST_STATUS));
            if (firstStatus.Text != status)
            {
                return false;
            }
            return true;
        }
        public void DuplicateSchedulerModalPage()
        {
            ShowPlusScheduler();
            var duplicateBtn = WaitForElementIsVisible(By.XPath(DUPLICATE_SCHEDULER));
            duplicateBtn.Click();
            WaitForLoad();
        }
        public void DuplicateSchedule(string schedulerName, string task, string newSchedulerName)
        {
            _scheduler = WaitForElementIsVisible(By.Id(SCHEDULER));
            _scheduler.SetValue(ControlType.DropDownList, schedulerName);
            WaitForLoad();
            _task = WaitForElementIsVisible(By.Id(TASK));
            _task.SetValue(ControlType.DropDownList, task);
            WaitForLoad();
            _newSchedulerName = WaitForElementIsVisible(By.Id(NEW_SCHEDULER_NAME));
            _newSchedulerName.SetValue(ControlType.TextBox, newSchedulerName);
            WaitForLoad();
            var duplicateTaskBtn = WaitForElementIsVisible(By.Id(DUPLICATE_SCHEDULER_BUTTON));
            duplicateTaskBtn.Click();
        }

        public void ClickFirstLine()
        {
            _detail_Button = WaitForElementIsVisible(By.XPath(DETAIL_BUTTON), nameof(DETAIL_BUTTON));
            WaitPageLoading();
            _detail_Button.Click();
            WaitForLoad();
        }

        public string GetFirstLineName()
        {
            var firstLineName = WaitForElementExists(By.XPath(FIRST_SCHEDULER_NAME));
            return firstLineName.Text.Trim();
        }

        public ScheduleDetailsPage EditScheduleDetailsPage()
        {
            ClickFirstLine();
            WaitForLoad();

            return new ScheduleDetailsPage(_webDriver, _testContext);
        }

        public ScheduleDetailsPage SelectFirstRow()
        {
            WaitForLoad();
            var firstRow = WaitForElementIsVisible(By.XPath(FIRST_ROW));
            firstRow.Click();
            WaitPageLoading();
            return new ScheduleDetailsPage(_webDriver, _testContext);
        }
        public void FillFields_EditSite(string site)
        {
            ComboBoxSelectById(new ComboBoxOptions(SITES, (string)site));
        }
        public string GetFirstLineSite()
        {
            if (IsDev())
            {
                var firstLineSite = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[4]"));
                WaitForLoad();
                return firstLineSite.Text;
            }
            else
            {
                var firstLineSite = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[4]"));
                return firstLineSite.Text;
            }
        }

        public bool CheckSites(string site)
        {
            _siteFirstRow = WaitForElementIsVisible(By.XPath(SITE_FIRST_ROW));
            return _siteFirstRow.Text.Contains(site);
        }

        public string GetFirstName()
        {
            var name = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[2]"));
            WaitForLoad();
            return name.Text;
        }

        public string GetFirstTaskName()
        {
            var taskName = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[3]"));
            WaitForLoad();
            return taskName.Text;
        }

        public string GetFirstSite()
        {
            var site = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[4]"));
            WaitForLoad();
            return site.Text;
        }

        public string GetFirstStatus()
        {
            var status = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[5]"));
            WaitForLoad();
            return status.Text;
        }

        public int GetNumberOfSchedulers()
        {
            IWebElement list = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody"));
            IList<IWebElement> rows = list.FindElements(By.XPath(".//tr"));
            int numberOfRows = rows.Count;
            return numberOfRows;
        }
    }
}
