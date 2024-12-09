using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Tasks;
using Newrest.Winrest.FunctionalTests.PageObjects;
using System;
using Newrest.Winrest.FunctionalTests.Utils;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Scheduler;
using Newrest.Winrest.FunctionalTests.PageObjects.TabletApp;

namespace Newrest.Winrest.FunctionalTests.TabletApp
{
    [TestClass]
    public class ToDoListTest : TestBase
    {
        private readonly string TASK_NAME = "Task-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private const int Timeout = 600000;

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_TODO_vérifToDoListRepeatTime()
        {
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string scheduleName = "NewSchedule_-" + DateUtils.Now.ToString("dd/MM/yyyy") + "_" + new Random().Next();
            string taskName = "NewTask_-" + DateUtils.Now.ToString("dd/MM/yyyy") + "_" + new Random().Next();
            string statusProduction = TestContext.Properties["StatusProduction"].ToString();
            HomePage homePage = LogInAsAdmin();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            tasksPage.CreateNewTask(taskName, siteACE, customer);
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            try
            {
                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewScheduleForTabletApp(scheduleName, taskName, siteACE, DateUtils.Now, statusProduction, DateUtils.Now.AddDays(-2), DateUtils.Now.AddDays(2), "6", "0600AM", "1159PM");
                homePage.Navigate();
                homePage.GotoTabletApp();
                ToDoListTabletAppPage toDoList = homePage.GotoTabletApp_ToDoList();
                toDoList.WaitLoading();
                bool exists = toDoList.VerifyAllHoursExist(taskName);
                Assert.IsTrue(exists, "La tâche ne se répète pas à 6h");
                toDoList.WaitPageLoading();
            }
            finally
            {
                ToDoListTabletAppPage tabletApp = new ToDoListTabletAppPage(WebDriver, TestContext);
                tabletApp.WaitPageLoading();
                tabletApp.GoToWinrestHome();
                homePage.GoToDoList_Scheduler();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                schedulerPage.DeleteScheduler();
                homePage.GoToDoList_Tasks();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_TODO_vérifToDoListDays()
        {
            string name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + new Random().Next();
            string site = "ACE";
            string customer = "3D COMM";
            string schelduerName = "Schedule" + "- For Test DUPLICATE" + DateUtils.Now.ToString("dd/MM/yyyy") + "_" + new Random().Next();
            string statusProduction = TestContext.Properties["StatusProduction"].ToString();
            HomePage homePage = LogInAsAdmin();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            try
            {
                tasksPage.ResetFilters();
                tasksPage.CreateNewTask(name, site, customer);
                TasksGeneralInfo tasksGeneralInfo = new TasksGeneralInfo(WebDriver, TestContext);
                tasksGeneralInfo.ClickOnSchedulersTab();
                tasksGeneralInfo.ShowPlusScheduler();
                tasksGeneralInfo.CreateSchedulerModalPage();
                System.Collections.Generic.List<string> plannedDays = tasksGeneralInfo.CreateNewSchedule(schelduerName, name, site, DateUtils.Now, statusProduction,
                        DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(30), "1", true, false, false, false, false, false, false);
                tasksGeneralInfo.BackToList();
                ToDoListTabletAppPage tabletApp = new ToDoListTabletAppPage(WebDriver, TestContext);
                tabletApp.GoToWinrestHome();
                tabletApp.GotoTabletApp();
                tabletApp.GotoTabletApp_ToDoList();
                TabletAppPage tabletAppPage = new TabletAppPage(WebDriver, TestContext);
                DateTime today = DateTime.Now;
                int daysUntilNext = ((int)DayOfWeek.Monday - (int)today.DayOfWeek + 7) % 7;
                DateTime nextMonday = today.AddDays(daysUntilNext);
                tabletAppPage.SelectDayInToDoList(nextMonday);
                System.Collections.Generic.List<string> taskNames = tabletAppPage.GetTaskNames();
                bool taskNameExistNextMonday = taskNames.Contains(name);
                Assert.IsTrue(taskNameExistNextMonday, $"La tâche '{name}' ne s'affiche pas le jour planifié {nextMonday}.");
                tabletAppPage.SelectDayInToDoList(nextMonday.AddDays(1));
                taskNames = tabletAppPage.GetTaskNames();
                bool taskNameNotExistNextTuesDay = !taskNames.Contains(name);
                Assert.IsTrue(taskNameNotExistNextTuesDay, $"La tâche '{name}'  s'affiche le jour non planifié {nextMonday}.");
            }
            finally
            {
                ToDoListTabletAppPage tabletApp = new ToDoListTabletAppPage(WebDriver, TestContext);
                tabletApp.WaitPageLoading();
                tabletApp.GoToWinrestHome();
                SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                schedulerPage.DeleteScheduler();
                homePage.GoToDoList_Tasks();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, name);
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();
            }
        }
    }
}
