using DocumentFormat.OpenXml.Office2013.WebExtentionPane;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Schedule;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Scheduler;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Tasks;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Policy;
using System.Threading;
using System.Windows.Forms;
namespace Newrest.Winrest.FunctionalTests.ToDoList
{
    [TestClass]
    public class SchedulerTest : TestBase
    {
        private static Random random = new Random();
        string SCHEDULE_NAME = "";
        string TASK_NAME = "";
        private const int Timeout = 600000;
        private const string Default_Time = "00:00:00.0000000";
        private const int STATUS_Test = 0;
        private const int STATUS_Production = 1;
        private const int STATUS_Obsolete = 2;
   
        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(TD_SC_Filter_SearchScheduler):
                    TestInitialize_CreateTaskForScheduler();
                    TestInitialize_CreateScheduler();
                    break;

                default:
                    break;
            }
        }

        [TestCleanup]
        public override void TestCleanup()
        {
            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(TD_SC_Filter_SearchScheduler):
                    TestCleanup_DeleteTaskForScheduler();
                    TestCleanup_DeleteScheduler();
                    break;

                default:
                    break;
            }
        }

        public void TestCleanup_DeleteTaskForScheduler()
        {
            DeleteTaskDefinition((int)TestContext.Properties[string.Format("TaskDefinitionId")]);

            Assert.IsNull(TaskDefinitionToSiteExist((int)TestContext.Properties[string.Format("TaskDefinitionId")]), "La Task n'est pas supprimé.");
        }

        public void TestCleanup_DeleteScheduler()
        {
            DeleteTaskScheduler((int)TestContext.Properties[string.Format("TaskSchedulerId")]);

            Assert.IsNull(TaskDefinitionToSiteExist((int)TestContext.Properties[string.Format("TaskSchedulerId")]), "Le Scheduler n'est pas supprimé.");
        }

        public void TestInitialize_CreateTaskForScheduler()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            TASK_NAME = "Task_" + DateUtils.Now.ToString("dd/MM/yyyy") + "_" + random.Next(100, 9999);

            TestContext.Properties[string.Format("TaskDefinitionId")] = InsertTaskDefinition(customer, TASK_NAME);
            InsertTaskDefinitionToSites(site, (int)TestContext.Properties[string.Format("TaskDefinitionId")]);

            Assert.IsNotNull(TaskDefinitionToSiteExist((int)TestContext.Properties[string.Format("TaskDefinitionId")]), "La Task n'est pas crée.");
        }

        public void TestInitialize_CreateScheduler()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            SCHEDULE_NAME = "scheduler_" + DateUtils.Now.ToString("dd/MM/yyyy") + "_" + random.Next(100, 9999);

            TestContext.Properties[string.Format("TaskSchedulerId")] = InsertTaskScheduler(SCHEDULE_NAME, STATUS_Test, DateTime.Today.AddDays(-6), DateTime.Today.AddDays(6), (int)TestContext.Properties[string.Format("TaskDefinitionId")]);
            InsertTaskSchedulerToSites(site, (int)TestContext.Properties[string.Format("TaskSchedulerId")]);

            Assert.IsNotNull(TaskSchedulerToSiteExist((int)TestContext.Properties[string.Format("TaskSchedulerId")]), "Le Scheduler n'est pas crée");
        }


        [Priority(0)]
        [TestMethod]
        [Timeout(TestTimeout.Infinite)]
        [TestCategory("Prerequis")]
        public void TD_SC_CreateData()
        {
            var statusProduction = TestContext.Properties["StatusProduction"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            schedulerPage.ResetFilters();
            schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, SCHEDULE_NAME);
            if (schedulerPage.CheckTotalNumber() > 0)
            {
                schedulerPage.DeleteScheduler();
            }
            TasksPage tasksPage = schedulerPage.GoToTaskTab();
            tasksPage.ResetFilters();
            tasksPage.Filter(TasksPage.FilterType.Search, TASK_NAME);
            if (tasksPage.CheckTotalNumber() == 0)
            {
                //Create Task
                tasksPage.CreateNewTaskWithMultipleSite(TASK_NAME, new List<string> { siteACE, siteMAD }, customer);
            }
            schedulerPage.GoToSchedulerTab();
            //Create Scheduler
            SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
            schedulerCreateModalPage.CreateNewSchedule(SCHEDULE_NAME, TASK_NAME, siteACE, DateUtils.Now, statusProduction,
                    DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(1), "2");
            schedulerPage.BackToList();
            schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, SCHEDULE_NAME);
            Assert.IsTrue(schedulerPage.CheckTotalNumber() != 0, $"Le scheduler {SCHEDULE_NAME} n'est pas ajoutée");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_Create_NewScheduler()
        {
            var scheduleName = "NewSchedule_" + DateUtils.Now.ToString("dd/MM/yyyy");
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            var statusProduction = TestContext.Properties["StatusProduction"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();
            ClearCache();

            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();

            schedulerPage.ResetFilters();
            schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);

            if (schedulerPage.CheckTotalNumber() > 0)
            {
                schedulerPage.DeleteScheduler();
            }

            ClearCache();
            schedulerPage = homePage.GoToDoList_Scheduler();
            try
            {
                // Création d'un nouveau Scheduler
                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(scheduleName, TASK_NAME, siteACE, DateUtils.Now, statusProduction,
                    DateUtils.Now.AddDays(-2), DateUtils.Now.AddDays(2), "2");

                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
                Assert.IsTrue(schedulerPage.CheckTotalNumber() != 0, $"Le scheduler {scheduleName} n'a pas été ajouté.");
            }
            finally
            {
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
                //Delete Scheduler ajouté
                schedulerPage.DeleteScheduler();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_Delete_Scheduler()
        {
            var scheduleName = "NewSchedule4Delete" + DateUtils.Now.ToString("dd/MM/yyyy");
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            var statusProduction = TestContext.Properties["StatusProduction"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            ClearCache();

            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            try
            {
                schedulerPage.ResetFilters();

                // Création du scheduler
                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(scheduleName, TASK_NAME, siteACE, DateUtils.Now, statusProduction,
                    DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(1), "2");

                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
                // Suppression du scheduler
                Assert.IsTrue(schedulerPage.CheckTotalNumber() > 0, "Le scheduler à supprimer est introuvable.");
                schedulerPage.DeleteScheduler();
            }
            finally
            {
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
                bool isSchedulerDeleted = schedulerPage.CheckTotalNumber() == 0;
                Assert.IsTrue(isSchedulerDeleted, $"Le scheduler '{scheduleName}' n'a pas été supprimé.");
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_Filter_SearchScheduler()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            
            schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, SCHEDULE_NAME); 
            schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
            bool isFilterCorrect = schedulerPage.VerifySchedulerNamesMatch(SCHEDULE_NAME);
            Assert.IsTrue(isFilterCorrect, "le filtre par  Scheduler Name ne fonctionne pas");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_Filter_SearchTaskName()
        {
            string taskName = TASK_NAME + "-" + new Random().Next().ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string schelduerName = SCHEDULE_NAME + "-" + new Random().Next().ToString();
            var statusProduction = TestContext.Properties["StatusProduction"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            ClearCache();
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            schedulerPage.ResetFilters();
            TasksPage tasksPage = schedulerPage.GoToTaskTab();
            tasksPage.ResetFilters();
            try
            {
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                if (tasksPage.CheckTotalNumber() == 0)
                {
                    //Create Task
                    tasksPage.CreateNewTaskWithMultipleSite(taskName, new List<string> { siteACE, siteMAD }, customer);
                }
                schedulerPage.GoToSchedulerTab();

                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(schelduerName, taskName, siteACE, DateUtils.Now, statusProduction,
                        DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(1), "2");
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();

                schedulerPage.Filter(SchedulerPage.FilterType.SearchTaskName, taskName);
                Assert.IsTrue(schedulerPage.VerifySearchTaskNameFilter(taskName), "le filtre par Task Name ne fonctionne pas");
            }
            finally
            {
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                //Delete Scheduler ajouté
                schedulerPage.DeleteScheduler();
                schedulerPage.ResetFilters();
                //Delete Task
                tasksPage = schedulerPage.GoToTaskTab();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();
            }



        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_Filter_Sites()
        {
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            ClearCache();
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            //Site Filter
            schedulerPage.Filter(SchedulerPage.FilterType.Sites, siteACE);
            bool isSiteFiltred = schedulerPage.VerifySiteFilter(siteACE);
            Assert.IsTrue(isSiteFiltred, "Le filtre par Site ne fonctionne pas");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_Filter_StatutTest()
        {
            string statusProduction = TestContext.Properties["StatusProduction"].ToString();
            string statusTest = TestContext.Properties["StatusTest"].ToString();
            string statusObsolete = "Obsolete";
            string prefix = DateTime.Now.ToString() + new Random().Next().ToString();
            string name1 = prefix + "Schelude1";
            string name2 = prefix + "Schelude2";
            string name3 = prefix + "Schelude3";
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            DateTime validToInput = DateUtils.Now.AddDays(2);
            DateTime validFromInput = DateUtils.Now.AddDays(-2);
            DateTime realisationDateInput = DateUtils.Now;
            string repeatTimeInput = "2";
            var filter = true;
            string taskName = "Test" + DateTime.Now;
            string customer = TestContext.Properties["CustomerLpFilter"].ToString();
            int check = 1;
            //Arrange
            var homePage = LogInAsAdmin();
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            // act
            try
            {
                // create a task
                TasksPage taskPage = homePage.GoToDoList_Tasks();
                taskPage.CreateNewTask(taskName, siteACE, customer);
                schedulerPage = taskPage.GoToDoList_Scheduler();
                //create Schelde1
                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(name1, taskName, siteACE, realisationDateInput, statusProduction,
                validFromInput, validToInput, repeatTimeInput);
                schedulerPage.BackToList();
                //create Schelde2
                schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(name2, taskName, siteACE, realisationDateInput, statusTest,
                validFromInput, validToInput, repeatTimeInput);
                schedulerPage.BackToList();
                //create Schelde3
                schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(name3, taskName, siteACE, realisationDateInput, statusObsolete,
                validFromInput, validToInput, repeatTimeInput);
                schedulerPage.BackToList();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, prefix);
                schedulerPage.WaitPageLoading();
                schedulerPage.Filter(SchedulerPage.FilterType.Test, filter);
                schedulerPage.WaitPageLoading();
                int count = schedulerPage.GetNumberOfSchedulers();
                Assert.AreEqual(check, count, "le filre sur status test ne s'applique pas");
                string status = schedulerPage.GetFirstStatus();
                Assert.AreEqual(status, statusTest, "le filre sur status test ne s'applique pas");

            }
            finally
            {
                // deleteTask
                var tasksPage = schedulerPage.GoToTaskTab();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.DeleteTask();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_Filter_StatutProduction()
        {
            string statusProduction = TestContext.Properties["StatusProduction"].ToString();
            string statusTest = TestContext.Properties["StatusTest"].ToString();
            string statusObsolete = "Obsolete";
            string prefix = DateTime.Now.ToString()+ new Random().Next().ToString();
            string name1 = prefix + "Schelude1";
            string name2 = prefix + "Schelude2";
            string name3 = prefix + "Schelude3";
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            DateTime validToInput = DateUtils.Now.AddDays(2);
            DateTime validFromInput = DateUtils.Now.AddDays(-2);
            DateTime realisationDateInput = DateUtils.Now;
            string repeatTimeInput = "2";
            var filter = true;
            string taskName = "Test" + DateTime.Now;
            string customer = TestContext.Properties["CustomerLpFilter"].ToString();
            int check = 1;
            //Arrange
            var homePage = LogInAsAdmin();
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            // act
            try
            {
                // create a task
                TasksPage taskPage = homePage.GoToDoList_Tasks();
                taskPage.CreateNewTask(taskName, siteACE, customer);
                schedulerPage = taskPage.GoToDoList_Scheduler();
                //create Schelde1
                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(name1, taskName, siteACE, realisationDateInput, statusProduction,
                validFromInput, validToInput, repeatTimeInput);
                schedulerPage.BackToList();
                //create Schelde2
                schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(name2, taskName, siteACE, realisationDateInput, statusTest,
                validFromInput, validToInput, repeatTimeInput);
                schedulerPage.BackToList();
                //create Schelde3
                schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(name3, taskName, siteACE, realisationDateInput, statusObsolete,
                validFromInput, validToInput, repeatTimeInput);
                schedulerPage.BackToList();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, prefix);
                schedulerPage.WaitPageLoading();
                schedulerPage.Filter(SchedulerPage.FilterType.Production, filter);
                schedulerPage.WaitPageLoading();
                int count = schedulerPage.GetNumberOfSchedulers();
                Assert.AreEqual(check, count, "le filre sur status production ne s'applique pas");
                string status = schedulerPage.GetFirstStatus();
                Assert.AreEqual(status, statusProduction, "le filre sur status production ne s'applique pas");

            }
            finally
            {
                // deleteTask
                var tasksPage = schedulerPage.GoToTaskTab();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.DeleteTask();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_Filter_StatutObsolete()
        {
            string statusProduction = TestContext.Properties["StatusProduction"].ToString();
            string statusTest = TestContext.Properties["StatusTest"].ToString();
            string statusObsolete = "Obsolete";
            string prefix = DateTime.Now.ToString() + new Random().Next().ToString();
            string name1 = prefix + "Schelude1";
            string name2 = prefix + "Schelude2";
            string name3 = prefix + "Schelude3";
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            DateTime validToInput = DateUtils.Now.AddDays(2);
            DateTime validFromInput = DateUtils.Now.AddDays(-2);
            DateTime realisationDateInput = DateUtils.Now;
            string repeatTimeInput = "2";
            var filter = true;
            string taskName = "Test" + DateTime.Now;
            string customer = TestContext.Properties["CustomerLpFilter"].ToString();
            int check = 1;
            //Arrange
            var homePage = LogInAsAdmin();
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            // act
            try
            {
                // create a task
                TasksPage taskPage = homePage.GoToDoList_Tasks();
                taskPage.CreateNewTask(taskName, siteACE, customer);
                schedulerPage = taskPage.GoToDoList_Scheduler();
                //create Schelde1
                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(name1, taskName, siteACE, realisationDateInput, statusProduction,
                validFromInput, validToInput, repeatTimeInput);
                schedulerPage.BackToList();
                //create Schelde2
                schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(name2, taskName, siteACE, realisationDateInput, statusTest,
                validFromInput, validToInput, repeatTimeInput);
                schedulerPage.BackToList();
                //create Schelde3
                schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(name3, taskName, siteACE, realisationDateInput, statusObsolete,
                validFromInput, validToInput, repeatTimeInput);
                schedulerPage.BackToList();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, prefix);
                schedulerPage.WaitPageLoading();
                schedulerPage.Filter(SchedulerPage.FilterType.Obsolete, filter);
                schedulerPage.WaitPageLoading();
                int count = schedulerPage.GetNumberOfSchedulers();
                Assert.AreEqual(check, count, "le filre sur status Obsolete ne s'applique pas");
                string status = schedulerPage.GetFirstStatus();
                Assert.AreEqual(status, statusObsolete, "le filre sur status Obsolete ne s'applique pas");
            }
            finally
            {
                // deleteTask
                var tasksPage = schedulerPage.GoToTaskTab();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.DeleteTask();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_ResetFilter()
        {
            object value = null;

            //Arrange
            var homePage = LogInAsAdmin();

            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            var numberDefaultListSiteFilter = schedulerPage.GetNumberSelectedSiteFilter();
            //Filter
            schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, "Test");
            schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
            schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
            schedulerPage.Filter(SchedulerPage.FilterType.Obsolete, true);
            schedulerPage.Filter(SchedulerPage.FilterType.SearchTaskName, "Test");
            schedulerPage.Filter(SchedulerPage.FilterType.From, DateTime.Now.AddDays(-1));
            schedulerPage.Filter(SchedulerPage.FilterType.To, DateTime.Now.AddDays(1));
            schedulerPage.Filter(SchedulerPage.FilterType.Sites, "ACE");
            //Reset filter
            schedulerPage.ResetFilters();
            //Verify reset filter
            value = schedulerPage.GetFilterValue(SchedulerPage.FilterType.SearchSchedulerName);
            Assert.AreEqual("", value, "ResetFilter Search SchedulerName ''");
            value = schedulerPage.GetFilterValue(SchedulerPage.FilterType.SearchTaskName);
            Assert.AreEqual("", value, " ResetFilter Search TaskName ''");
            value = schedulerPage.GetFilterValue(SchedulerPage.FilterType.Test);
            Assert.AreEqual(false, value, "ResetFilter Test");
            value = schedulerPage.GetFilterValue(SchedulerPage.FilterType.Production);
            Assert.AreEqual(true, value, "ResetFilter Production");
            value = schedulerPage.GetFilterValue(SchedulerPage.FilterType.Obsolete);
            Assert.AreEqual(false, value, "ResetFilter Obsolete");
            var numberSelectedSiteFilter = schedulerPage.GetNumberSelectedSiteFilter();
            Assert.AreEqual(numberSelectedSiteFilter, numberDefaultListSiteFilter, "ResetFilter Site ''");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_Duplicate_Scheduler()
        {
            var schedulernameDuplicate = "Duplicate Schedule" + DateUtils.Now.ToString("dd/MM/yyyy"); ;

            string siteACE = TestContext.Properties["SiteACE"].ToString();
            var statusProduction = TestContext.Properties["StatusProduction"].ToString();
            string schelduerName = SCHEDULE_NAME + "-" + new Random().Next().ToString();
            string taskName = TASK_NAME + "-" + new Random().Next().ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            ClearCache();
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            TasksPage tasksPage = schedulerPage.GoToTaskTab();
            tasksPage.ResetFilters();
            tasksPage.Filter(TasksPage.FilterType.Search, taskName);
            if (tasksPage.CheckTotalNumber() == 0)
            {
                //Create Task
                tasksPage.CreateNewTaskWithMultipleSite(taskName, new List<string> { siteACE, siteMAD }, customer);
            }
            schedulerPage.GoToSchedulerTab();
            SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
            schedulerCreateModalPage.CreateNewSchedule(schelduerName, taskName, siteACE, DateUtils.Now, statusProduction,
                    DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(1), "2");
            schedulerPage.BackToList();
            schedulerPage.ResetFilters();
            try
            {
                //Duplicate Scheduler
                schedulerPage.DuplicateSchedulerModalPage();
                schedulerPage.DuplicateSchedule(schelduerName, taskName, schedulernameDuplicate);
                schedulerPage.BackToList();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schedulernameDuplicate);
                schedulerPage.Filter(SchedulerPage.FilterType.SearchTaskName, taskName);
                Assert.IsTrue(schedulerPage.VerifySearchSchedulerFilter(schedulernameDuplicate) && schedulerPage.VerifySearchTaskNameFilter(taskName), "La duplication ne fonctionne pas.");
            }
            finally
            {
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchTaskName, taskName);
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schedulernameDuplicate);
                schedulerPage.DeleteScheduler();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                schedulerPage.DeleteScheduler();
                schedulerPage.GoToTaskTab();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.DeleteTask();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_Filter_DateFrom()
        {
            // Prepare
            Random r = new Random();
            string nameDemain = "Scheduler_demain_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            string namePlage = "Scheduler_plage_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            string nameAuj = "Scheduler_auj_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            string site = TestContext.Properties["SiteACE"].ToString();
            string customer = TestContext.Properties["CustomerLpFilter"].ToString();
            string status = TestContext.Properties["StatusTest"].ToString();
            string taskName = "Test" + DateTime.Now ;

            //Arrange
            var homePage = LogInAsAdmin();

            ClearCache();
            try
            {
                //create task
                TasksPage taskPage = homePage.GoToDoList_Tasks();
                taskPage.CreateNewTask(taskName, site, customer);

                //create scheduler avant fromDate
                SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
                SchedulerCreateModalPage modal = schedulerPage.CreateSchedulerModalPage();
                modal.CreateNewSchedule(namePlage, taskName, site, DateUtils.Now.AddDays(-2), status, DateUtils.Now.AddDays(-3), DateUtils.Now.AddDays(-1), "3");

                //create scheduler égale à fromDate
                schedulerPage.BackToList();
                modal = schedulerPage.CreateSchedulerModalPage();
                modal.CreateNewSchedule(nameAuj, taskName, site, DateUtils.Now.AddDays(2), status, DateUtils.Now, DateUtils.Now.AddDays(5), "3");

                //create scheduler aprés fromDate
                schedulerPage.BackToList();
                modal = schedulerPage.CreateSchedulerModalPage();
                modal.CreateNewSchedule(nameDemain, taskName, site, DateUtils.Now.AddDays(2), status, DateUtils.Now.AddDays(1), DateUtils.Now.AddDays(5), "3");
                schedulerPage.BackToList();

                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.Filter(SchedulerPage.FilterType.From, DateUtils.Now);
                schedulerPage.Filter(SchedulerPage.FilterType.To, DateUtils.Now.AddDays(5));
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, nameDemain);

                int totalNumber = schedulerPage.CheckTotalNumber();
                Assert.AreEqual(1, totalNumber, "Le filtre From ne fonctionne pas correctement.");

                // Appliquer le filtre avant Date From
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.Filter(SchedulerPage.FilterType.From, DateUtils.Now);
                schedulerPage.Filter(SchedulerPage.FilterType.To, DateUtils.Now.AddDays(5));
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, namePlage);

                totalNumber = schedulerPage.CheckTotalNumber();
                // Les données sont filtré et correspondent
                Assert.AreEqual(0,totalNumber , "Le filtre From ne fonctionne pas correctement.");

                // Appliquer le filtre en Date From
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.Filter(SchedulerPage.FilterType.From, DateUtils.Now);
                schedulerPage.Filter(SchedulerPage.FilterType.To, DateUtils.Now.AddDays(5));
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, nameAuj);

                totalNumber = schedulerPage.CheckTotalNumber();
                Assert.AreEqual(1, totalNumber, "Le filtre From ne fonctionne pas correctement.");
            }
            finally
            {
                SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();

                //Delete Scheduler
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, nameDemain);
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.DeleteScheduler();

                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, namePlage);
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.DeleteScheduler();

                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, nameAuj);
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.DeleteScheduler();

                //Delete Task
                var taskPage = homePage.GoToDoList_Tasks();
                taskPage.Filter(TasksPage.FilterType.Search, taskName);
                taskPage.DeleteTask();
            }
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_Filter_DateTo()
        {
            // Prepare
            Random r = new Random();
            string nameDemain = "Scheduler_demain_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            string namePlage = "Scheduler_plage_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            string nameAuj = "Scheduler_auj_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            string site = TestContext.Properties["SiteACE"].ToString();
            string customer = TestContext.Properties["CustomerLpFilter"].ToString();
            string status = TestContext.Properties["StatusTest"].ToString();
            string taskName = "Test" + DateTime.Now;

            //Arrange
            var homePage = LogInAsAdmin();

            ClearCache();
            try
            {
                //create task
                TasksPage taskPage = homePage.GoToDoList_Tasks();
                taskPage.CreateNewTask(taskName, site, customer);

                //create scheduler avant fromDate
                SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
                SchedulerCreateModalPage modal = schedulerPage.CreateSchedulerModalPage();
                modal.CreateNewSchedule(namePlage, taskName, site, DateUtils.Now.AddDays(2), status, DateUtils.Now, DateUtils.Now.AddDays(3), "3");

                //create scheduler égale à fromDate
                schedulerPage.BackToList();
                modal = schedulerPage.CreateSchedulerModalPage();
                modal.CreateNewSchedule(nameAuj, taskName, site, DateUtils.Now.AddDays(2), status, DateUtils.Now, DateUtils.Now.AddDays(5), "3");

                //create scheduler aprés fromDate
                schedulerPage.BackToList();
                modal = schedulerPage.CreateSchedulerModalPage();
                modal.CreateNewSchedule(nameDemain, taskName, site, DateUtils.Now.AddDays(2), status, DateUtils.Now, DateUtils.Now.AddDays(8), "3");
                schedulerPage.BackToList();

                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.Filter(SchedulerPage.FilterType.From, DateUtils.Now);
                schedulerPage.Filter(SchedulerPage.FilterType.To, DateUtils.Now.AddDays(5));
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, nameDemain);

                int totalNumber = schedulerPage.CheckTotalNumber();
                Assert.AreEqual(1, totalNumber, "Le filtre To ne fonctionne pas correctement.");

                // Appliquer le filtre avant Date From
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.Filter(SchedulerPage.FilterType.From, DateUtils.Now);
                schedulerPage.Filter(SchedulerPage.FilterType.To, DateUtils.Now.AddDays(5));
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, namePlage);

                totalNumber = schedulerPage.CheckTotalNumber();
                // Les données sont filtré et correspondent
                Assert.AreEqual(1, totalNumber, "Le filtre To ne fonctionne pas correctement.");

                // Appliquer le filtre en Date From
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.Filter(SchedulerPage.FilterType.From, DateUtils.Now);
                schedulerPage.Filter(SchedulerPage.FilterType.To, DateUtils.Now.AddDays(5));
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, nameAuj);

                totalNumber = schedulerPage.CheckTotalNumber();
                Assert.AreEqual(1, totalNumber, "Le filtre To ne fonctionne pas correctement.");
            }
            finally
            {
                SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();

                //Delete Scheduler
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, nameDemain);
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.DeleteScheduler();

                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, namePlage);
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.DeleteScheduler();

                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, nameAuj);
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.DeleteScheduler();

                //Delete Task
                var taskPage = homePage.GoToDoList_Tasks();
                taskPage.Filter(TasksPage.FilterType.Search, taskName);
                taskPage.DeleteTask();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_editName()
        {
            var updated_Name = "Test_Updated";
            //Arrange
            var homePage = LogInAsAdmin();

            ClearCache();
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();

            try
            {
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, SCHEDULE_NAME);
                ScheduleDetailsPage scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
                scheduleDetailsPage.EditName(updated_Name);
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, updated_Name);
                var firstName = schedulerPage.GetFirstLineName();
                Assert.AreEqual(firstName, updated_Name, "Scheduler Name n'est pas modifieé");

            }
            finally
            {
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, updated_Name);
                ScheduleDetailsPage scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
                scheduleDetailsPage.EditName(SCHEDULE_NAME);
                schedulerPage.BackToList();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_editStartTime()
        {
            string schelduerName = "SchedulerForTest - " + new Random().Next().ToString();
            string taskName = "TaskForTest - " + new Random().Next().ToString();
            var statusProduction = TestContext.Properties["StatusProduction"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            ClearCache();

            //Act
            //Create Task
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            schedulerPage.ResetFilters();
            schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
            try
            {
                TasksPage tasksPage = schedulerPage.GoToTaskTab();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                //Create Task
                tasksPage.CreateNewTaskWithMultipleSite(taskName, new List<string> { siteACE, siteMAD }, customer);
                schedulerPage.GoToSchedulerTab();
                //Create Scheduler
                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(schelduerName, taskName, siteACE, DateUtils.Now, statusProduction,
                        DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(1), "2");
                
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                ScheduleDetailsPage scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
                var originStartTime = scheduleDetailsPage.GetStartTime();

                //Modifier Start Time
                scheduleDetailsPage.EditStartTime("0300AM");
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
                var editedStartTime = scheduleDetailsPage.GetStartTime();
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();

                //Assert
                Assert.AreNotEqual(originStartTime, editedStartTime, "Start time was not edited correctly.");

            }
            finally
            {
                //Supprimer Scheduler
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                schedulerPage.DeleteScheduler();
                //Supprimer Task
                TasksPage tasksPage = schedulerPage.GoToTaskTab();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_editEndTime()
        {
            string schelduerName = $"SchedulerForTest - {new Random().Next()}";
            string taskName = $"TaskForTest - { new Random().Next()}";
            string statusProduction = TestContext.Properties["StatusProduction"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();

            //Arrange
            PageObjects.HomePage homePage = LogInAsAdmin();

            ClearCache();

            //Act
            //Create Task
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            schedulerPage.ResetFilters();
            try
            {
                TasksPage tasksPage = schedulerPage.GoToTaskTab();
                tasksPage.ResetFilters();
                tasksPage.CreateNewTaskWithMultipleSite(taskName, new List<string> { siteACE, siteMAD }, customer);
                schedulerPage.GoToSchedulerTab();
                //Create Scheduler
                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(schelduerName, taskName, siteACE, DateUtils.Now, statusProduction,
                        DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(1), "2");

                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                ScheduleDetailsPage scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
                string originEndTime = scheduleDetailsPage.GetEndTime();

                scheduleDetailsPage.EditEndTime("0300AM");
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
                string editedEndTime = scheduleDetailsPage.GetEndTime();
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();

                //Assert
                Assert.AreNotEqual(originEndTime, editedEndTime, "End time was not edited correctly.");
            }
            finally
            {
                //Supprimer Scheduler
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                schedulerPage.DeleteScheduler();
                //Supprimer Task
                TasksPage tasksPage = schedulerPage.GoToTaskTab();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_editRealisation_Time()
        {
            string nameTask = "TaskTest";
            string name = "SchedulersTest";
            string site = "ACE";
            string status = "Production";


            //Arrange
            PageObjects.HomePage homePage = LogInAsAdmin();

            ClearCache();

            //Create Task
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            TasksPage tasksPage = schedulerPage.GoToTaskTab();
            try
            {
                tasksPage.CreateNewTask(nameTask, site, "3D COMM");
                schedulerPage = schedulerPage.GoToSchedulerTab();
                string numberOfSchedulerInHeader = schedulerPage.GetNumberOfSchedulerInHeader();
                //Create Scheduler
                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(name, nameTask, site, DateUtils.Now, status,
                    DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(1), "2");
                schedulerPage.BackToList();
                string numberOfSchedulerInHeaderAfterCreate = schedulerPage.GetNumberOfSchedulerInHeader();
                Assert.AreEqual(int.Parse(numberOfSchedulerInHeaderAfterCreate), int.Parse(numberOfSchedulerInHeader) + 1, "Scheduler n'est pas ajoutée");
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, name);
                ScheduleDetailsPage scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
                string realtionTimeBefore = scheduleDetailsPage.GetRealisationTime();
                scheduleDetailsPage.EditRealisationTime();
                string realtionTimeAfter = scheduleDetailsPage.GetRealisationTime();
                schedulerPage.BackToList();
                Assert.AreNotEqual(realtionTimeBefore, realtionTimeAfter, "Scheduler Realisation Time n'est pas modifieé");

            }
            finally
            {
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, name);
                //Delete Scheduler ajouté
                schedulerPage.DeleteScheduler();
                schedulerPage.ResetFilters();
                //Delete Task
                tasksPage = schedulerPage.GoToTaskTab();
                tasksPage.Filter(TasksPage.FilterType.Search, nameTask);
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_editStatus()
        {
            string statusProduction = TestContext.Properties["StatusProduction"].ToString();
            string updatedStatus = TestContext.Properties["StatusTest"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            schedulerPage.ResetFilters();
            schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, SCHEDULE_NAME);
            schedulerPage.Filter(SchedulerPage.FilterType.SearchTaskName, TASK_NAME);
            ScheduleDetailsPage scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
            try
            {
                scheduleDetailsPage.EditStatus(updatedStatus);
                schedulerPage = scheduleDetailsPage.BackToList();
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, SCHEDULE_NAME);
                bool isFilterCorrect = schedulerPage.VerifyStatusFilter(updatedStatus);
                Assert.IsTrue(isFilterCorrect, "Le filtre de statut ne fonctionne pas correctement.");
            }
            finally
            {
                scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
                scheduleDetailsPage.EditStatus(statusProduction);
                scheduleDetailsPage.BackToList();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_editRealisation_Date()
        {
            //prepare
            var updateRealisationDate = DateUtils.Now.AddDays(1);

            //arrange
            var homePage = LogInAsAdmin();

            ClearCache();

            //act
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            schedulerPage.ResetFilters();
            schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
            var scheduleName = schedulerPage.GetFirstName();
            schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
            ScheduleDetailsPage scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
            var realtionDateDefault = scheduleDetailsPage.GetRealisationDate();
            try
            {
                scheduleDetailsPage.EditRealisationDate(updateRealisationDate);
                schedulerPage = scheduleDetailsPage.BackToList();
                schedulerPage.Filter(SchedulerPage.FilterType.Test, true);
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
                Assert.AreEqual(updateRealisationDate.ToString("dd/MM/yyyy"), scheduleDetailsPage.GetRealisationDate(), "Scheduler Realisation Date n'est pas modifié");
            }
            finally
            {
                DateTime realtionDateConverted = DateTime.ParseExact(realtionDateDefault, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                scheduleDetailsPage.EditRealisationDate(realtionDateConverted);
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_editTask()
        {

            HomePage homePage = LogInAsAdmin();

            ClearCache();

            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            ScheduleDetailsPage scheduleDetailsPage = schedulerPage.EditScheduleDetailsPage();
            // Verify if task field is disabled
            bool isTaskFieldDisabled = scheduleDetailsPage.IsTaskFieldDisabled();
            Assert.IsTrue(isTaskFieldDisabled, "Le champ Task n'est pas grisé");

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_editSite()
        {
            // Prepare
            string siteValueToEdit = TestContext.Properties["SiteACE"].ToString();
            string oldSiteValue = "";
            // Arrange
            var homePage = LogInAsAdmin();

            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            schedulerPage.ResetFilters();
            schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
            try
            {
                // get first sechulder && edit site by a new value
                var scheduleName = schedulerPage.GetFirstLineName();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                oldSiteValue = schedulerPage.GetFirstLineSite();
                var firstItem = schedulerPage.SelectFirstRow();
                firstItem.EditSite(siteValueToEdit);
                schedulerPage = firstItem.BackToList();
                schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                var firstSite = schedulerPage.GetFirstLineSite();
                Assert.AreEqual(siteValueToEdit, firstSite, "Scheduler Site n'est pas modifieé");
            }
            finally
            {

                var firstItem = schedulerPage.SelectFirstRow();
                firstItem.EditSite(oldSiteValue);
                firstItem.BackToList();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_addSite()
        {
            var name = "NewSchedule";
            var nameTask = "NewTaskName";
            var status = "Test";
            List<string> sites = new List<string> { "ACE", "MAD" };

            var homePage = LogInAsAdmin();

            var toDoListTaskPage = homePage.GoToDoList_Tasks();
            toDoListTaskPage.CreateNewTaskWithMultipleSite(nameTask, sites, "3D COMM");
            var toDoListScheduler = homePage.GoToDoList_Scheduler();
            var createScheduler = toDoListScheduler.CreateSchedulerModalPage();
            ScheduleDetailsPage scheduleDetailPage = createScheduler.CreateNewSchedule(name, nameTask, "MAD", DateUtils.Now, status,
                DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(1), "2");
            scheduleDetailPage.EditSite("ACE");
            scheduleDetailPage.BackToList();
            toDoListScheduler.Filter(SchedulerPage.FilterType.Test, true);
            // trié du plus ancien au nouveau
            toDoListScheduler.PageSize("100");
            var exist = toDoListScheduler.CheckSites("ACE");
            Assert.IsTrue(exist, "Le changement n'est pas pris en compte");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_DuplicateSchedulerName()
        {
            var newSchedulerName = "Duplicate Scheduler_" + DateTime.Now.ToString("dd/MM/yyyy");
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            var statusProduction = TestContext.Properties["StatusProduction"].ToString();
            string schelduerName = SCHEDULE_NAME + "-" + new Random().Next().ToString();
            string taskName = TASK_NAME + "-" + new Random().Next().ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            ClearCache();

            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            schedulerPage.ResetFilters();


            TasksPage tasksPage = schedulerPage.GoToTaskTab();
            tasksPage.ResetFilters();
            tasksPage.Filter(TasksPage.FilterType.Search, taskName);
            if (tasksPage.CheckTotalNumber() == 0)
            {
                //Create Task
                tasksPage.CreateNewTaskWithMultipleSite(taskName, new List<string> { siteACE, siteMAD }, customer);
            }
            schedulerPage.GoToSchedulerTab();
            SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
            schedulerCreateModalPage.CreateNewSchedule(schelduerName, taskName, siteACE, DateUtils.Now, statusProduction,
                    DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(1), "2");
            schedulerPage.BackToList();
            schedulerPage.ResetFilters();

            schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
            var name = schedulerPage.GetFirstName();
            var nameTask = schedulerPage.GetFirstTaskName();
            try
            {

                schedulerPage.DuplicateSchedulerModalPage();
                schedulerPage.DuplicateSchedule(name, nameTask, newSchedulerName);
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchTaskName, nameTask);
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, newSchedulerName);
                var rowNumberNewScheduler = schedulerPage.CheckTotalNumber();
                Assert.AreNotEqual(0, rowNumberNewScheduler, "Aucun scheduler n'a été créé avec le nouveau nom.");
            }
            finally
            {
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchTaskName, nameTask);
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, newSchedulerName);
                schedulerPage.DeleteScheduler();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                schedulerPage.DeleteScheduler();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_editValidFromPeriod()
        {
            // prepare 
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string taskName = TASK_NAME + "-" + new Random().Next().ToString();
            var scheduleName = "NewSchedule_" + DateUtils.Now.ToString("dd/MM/yyyy") + new Random().Next().ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            var statusProduction = TestContext.Properties["StatusProduction"].ToString();
            var validFromDateEdited = DateUtils.Now.AddDays(-4);
            var validFromDate = DateUtils.Now.AddDays(-1);
            var validToDate = DateUtils.Now.AddDays(1);
            var realisationDateInput = DateUtils.Now;
            string repeatTimeInput = "2";
            //Arrange
            var homePage = LogInAsAdmin();
            // act
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            TasksPage tasksPage = schedulerPage.GoToTaskTab();
            tasksPage.ResetFilters();
            try
            {
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                if (tasksPage.CheckTotalNumber() == 0)
                {
                    //Create Task
                    tasksPage.CreateNewTaskWithMultipleSite(taskName, new List<string> { siteACE, siteMAD }, customer);
                }
                schedulerPage.GoToSchedulerTab();

                // create a Schelude 
                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(scheduleName, taskName, siteACE, realisationDateInput, statusProduction,
                validFromDate, validToDate, repeatTimeInput);
                // get the created Schelude
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
                var firstItem = schedulerPage.SelectFirstRow();
                // edit vadlidFrom
                firstItem.EditValidPeriod(validFromDateEdited, validToDate);
                // check
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
                var firstItemToVerif = schedulerPage.SelectFirstRow();
                var validToDateToVerif = firstItemToVerif.GePeriodFromDate();
                var validFromDateFormated = validFromDateEdited.ToString("dd/MM/yyyy");
                Assert.AreEqual(validToDateToVerif, validFromDateFormated, "le mise à jour ne s'applique pas sur Valid TO ");
            }
            finally
            {
                //Delete Schedule ajouté
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
                schedulerPage.DeleteScheduler();
                //Delete Task
                tasksPage = schedulerPage.GoToTaskTab();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.DeleteTask();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_editValidToPeriod()
        {
            //Prepare
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string taskName = TASK_NAME + "-" + new Random().Next().ToString();
            string scheduleName = "NewSchedule_" + DateUtils.Now.ToString("dd/MM/yyyy") + new Random().Next().ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string statusProduction = TestContext.Properties["StatusProduction"].ToString();
            DateTime validToDateEdited = DateUtils.Now.AddDays(4);
            DateTime validFromDate = DateUtils.Now.AddDays(-1);
            DateTime validToDate = DateUtils.Now.AddDays(1);
            DateTime realisationDateInput = DateUtils.Now;
            string repeatTimeInput = "2";

            //login
            PageObjects.HomePage homePage = LogInAsAdmin();
            SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
            TasksPage tasksPage = schedulerPage.GoToTaskTab();
            tasksPage.ResetFilters();
            try
            {
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                if (tasksPage.CheckTotalNumber() == 0)
                {
                    //Create Task
                    tasksPage.CreateNewTaskWithMultipleSite(taskName, new List<string> { siteACE, siteMAD }, customer);
                }
                schedulerPage.GoToSchedulerTab();

                // create a Schelude 
                SchedulerCreateModalPage schedulerCreateModalPage = schedulerPage.CreateSchedulerModalPage();
                schedulerCreateModalPage.CreateNewSchedule(scheduleName, taskName, siteACE, realisationDateInput, statusProduction, validFromDate, validToDate, repeatTimeInput);
                // get the created Schelude
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
                ScheduleDetailsPage firstItem = schedulerPage.SelectFirstRow();
                // edit validTo
                firstItem.EditValidPeriod(validFromDate, validToDateEdited);
                // check
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
                ScheduleDetailsPage firstItemToVerif = schedulerPage.SelectFirstRow();
                string validToDateToVerif = firstItemToVerif.GePeriodToDate();
                string validToDateFormatted = validToDateEdited.ToString("dd/MM/yyyy");
                Assert.AreEqual(validToDateToVerif, validToDateFormatted, "le mise à jour ne s'applique pas sur Valid TO ");
            }
            finally
            {
                //Delete Scheduler ajouté
                schedulerPage.BackToList();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, scheduleName);
                schedulerPage.Filter(SchedulerPage.FilterType.Production, true);
                schedulerPage.DeleteScheduler();
                //Delete Task
                tasksPage = schedulerPage.GoToTaskTab();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.DeleteTask();
            }
        }

        [Timeout(Timeout)]
        public void TD_SC_DuplicateTask()
        {
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var Newname = "NewTask_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = TestContext.Properties["SiteACE"].ToString();
            var customer = TestContext.Properties["CustomerJob"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            var tasksGeneralInfo = tasksPage.CreateNewTask(name, site, customer);
            string Name = tasksGeneralInfo.GetName();
            string SiteValue = tasksGeneralInfo.GetSiteValue();
            string CustomerValue = tasksGeneralInfo.GetCustomer();
            string VehiculRegistration = tasksGeneralInfo.GetVehiculeRegistration();
            tasksPage.BackToList();
            tasksPage.CreateFromOtherTasks(name, Newname, false);
            string NameAfterDuplicate = tasksGeneralInfo.GetName();
            string SiteValueAfterDuplicated = tasksGeneralInfo.GetSiteValue();
            string CustomerValueAfterDuplicated = tasksGeneralInfo.GetCustomer();
            string VehiculRegistrationAfterDuplicated = tasksGeneralInfo.GetVehiculeRegistration();

            Assert.AreNotEqual(NameAfterDuplicate, Name, "Le nom du Task est dupliqué");
            Assert.AreEqual(SiteValueAfterDuplicated, SiteValue, "Le site du Task dupliquée et celle de l'originale ne sont pas égales.");
            Assert.AreEqual(CustomerValueAfterDuplicated, CustomerValue, "Le customer du Task dupliquée et celle de l'originale ne sont pas égales.");
            Assert.AreEqual(VehiculRegistrationAfterDuplicated, VehiculRegistration, "La Registration du Task dupliquée et celle de l'originale ne sont pas égales.");

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_SC_DuplicateTaskAndScheduler()
        {
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var Newname = "NewTask_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = TestContext.Properties["SiteACE"].ToString();
            var customer = TestContext.Properties["CustomerJob"].ToString();
            string taskName = TASK_NAME + "- For Test DUPLICATE";
            string schelduerName = "Schedule" + "- For Test DUPLICATE";
            var statusProduction = TestContext.Properties["StatusProduction"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            var tasksGeneralInfo = tasksPage.CreateNewTask(name, site, customer);
            string Name = tasksGeneralInfo.GetName();
            string SiteValue = tasksGeneralInfo.GetSiteValue();
            string CustomerValue = tasksGeneralInfo.GetCustomer();
            string VehiculRegistration = tasksGeneralInfo.GetVehiculeRegistration();
            tasksGeneralInfo.ClickOnSchedulersTab();
            tasksGeneralInfo.CreateSchedulerModalPage();
            tasksGeneralInfo.CreateNewSchedule(schelduerName, name, site, DateUtils.Now, statusProduction, DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(1), "2");

            string SchedulersNameOnSchedulers = tasksGeneralInfo.GetNameOnSchedulers();
            string TaskNameOnSchedulers = tasksGeneralInfo.GetTaskNameOnSchedulers();
            string SitesOnSchedulers = tasksGeneralInfo.GetSiteOnSchedulers();
            string statutsOnSchedulers = tasksGeneralInfo.GetStatutOnSchedulers();

            tasksPage.BackToList();
            tasksPage.CreateFromOtherTasks(name, Newname, true);

            string NameAfterDuplicate = tasksGeneralInfo.GetName();
            string SiteValueAfterDuplicated = tasksGeneralInfo.GetSiteValue();
            string CustomerValueAfterDuplicated = tasksGeneralInfo.GetCustomer();
            string VehiculRegistrationAfterDuplicated = tasksGeneralInfo.GetVehiculeRegistration();
            tasksGeneralInfo.ClickOnSchedulersTab();
            string SchedulersNameOnSchedulersAfterDuplicate = tasksGeneralInfo.GetNameOnSchedulers();
            string TaskNameOnSchedulersAfterDuplicate = tasksGeneralInfo.GetTaskNameOnSchedulers();
            string SitesOnSchedulersAfterDuplicate = tasksGeneralInfo.GetSiteOnSchedulers();
            string statutsOnSchedulersAfterDuplicate = tasksGeneralInfo.GetStatutOnSchedulers();

            //Assert
            Assert.AreNotEqual(NameAfterDuplicate, Name, "Le nom du Task est dupliqué");
            Assert.AreNotEqual(TaskNameOnSchedulers, TaskNameOnSchedulersAfterDuplicate, "Le nom du Task on Schedule est dupliqué");
            Assert.AreEqual(SiteValueAfterDuplicated, SiteValue, "Le site du Task dupliquée et celle de l'originale ne sont pas égaux.");
            Assert.AreEqual(CustomerValueAfterDuplicated, CustomerValue, "Le customer du Task dupliquée et celle de l'originale ne sont pas égaux.");
            Assert.AreEqual(VehiculRegistrationAfterDuplicated, VehiculRegistration, "La Registration du Task dupliquée et celle de l'originale ne sont pas égales.");
            Assert.AreEqual(SchedulersNameOnSchedulers, SchedulersNameOnSchedulersAfterDuplicate, "Le nom du schedule  dupliquée et celle de l'originale ne sont pas égaux.");
            Assert.AreEqual(SitesOnSchedulers, SitesOnSchedulersAfterDuplicate, "Le site du schedule  dupliquée et celle de l'originale ne sont pas égaux.");
            Assert.AreEqual(statutsOnSchedulers, statutsOnSchedulersAfterDuplicate, "La status du schedule  dupliquée et celle de l'originale ne sont pas égales.");


        }
        ////////////////////////////////Insert Task for Scheduler////////////////////////////////

        private int InsertTaskDefinition(string customer, string name)
        {
            string query = @"
            DECLARE @customerId INT;
            SELECT TOP 1 @customerId = Id FROM customers WHERE Name LIKE @customer;

            Insert into [TaskDefinitions] ([Name]
              ,[CustomerId]
              ,[Category]
              ,[IsActive]
              ,[EnforceOrder]
              ,[ValidationRequired])
	          values
	          (
	          @name,
	          @customerId,
	          '0',
	          1,
	          0,
	          0
	          );
             SELECT SCOPE_IDENTITY();
            "
            ;
            return ExecuteAndGetInt(
            query,
                new KeyValuePair<string, object>("customer", customer),
                new KeyValuePair<string, object>("name", name)
            );
        }

        private void InsertTaskDefinitionToSites(string site, int taskDefId)
        {
            string query = @"
            DECLARE @siteId INT;
            SELECT TOP 1 @siteId = Id FROM sites WHERE Name LIKE @site;
 
            Insert into [dbo].[TaskDefinitionsToSites]
            (TaskDefinitionId
              ,[SiteId])
	          values
	          (@taskDefId, @siteId);
            "
            ;
            ExecuteNonQuery(
            query,
                new KeyValuePair<string, object>("site", site),
                new KeyValuePair<string, object>("taskDefId", taskDefId)
            );
        }

        private int? TaskDefinitionToSiteExist(int taskDefId)
        {
            string query = @"
             select TaskDefinitionId from TaskDefinitionsToSites where TaskDefinitionId = @taskDefId

              SELECT SCOPE_IDENTITY();
            "
            ;
            int? result = ExecuteAndGetInt(
            query,
            new KeyValuePair<string, object>("taskDefId", taskDefId)
            );
            return result == 0 ? (int?)null : result;
        }

        private void DeleteTaskDefinition(int taskDefId)
        {
            string query = @"
            delete from TaskDefinitions where Id = @taskDefId
            "
            ;
            ExecuteNonQuery(
            query,
                new KeyValuePair<string, object>("taskDefId", taskDefId)
            );
        }

        ////////////////////////////////Insert Scheduler////////////////////////////////

        private int InsertTaskScheduler(string name, int status, DateTime dateFrom, DateTime dateTo, int taskDefId, string timeFrom = Default_Time, string timeTo = Default_Time)
        {
            string query = @"
            Insert into [dbo].[TaskSchedulers] ([Name]
              ,[TaskDefinitionId]
              ,[RealisationDate]
              ,[Status]
              ,[DateFrom]
              ,[DateTo]
              ,[IsActiveOnMonday]
              ,[IsActiveOnTuesday]
              ,[IsActiveOnWednesday]
              ,[IsActiveOnThursday]
              ,[IsActiveOnFriday]
              ,[IsActiveOnSaturday]
              ,[IsActiveOnSunday]
              ,[TypeOfTimeRepeat]
              ,[TimeValueRepeat]
              ,[TimeFrom]
              ,[TimeTo])
	          values (
	          @name,
	          @taskDefId,
	           GETDATE(),
	          @status,
	          @dateFrom,
	          @dateTo,
	          0,
	          0,
	          0,
	          0,
	          0,
	          0,
	          0,
	          '0',
	          '0',
	          @timeFrom,
	          @timeTo
	          );
             SELECT SCOPE_IDENTITY();
            "
            ;
            return ExecuteAndGetInt(
            query,
                new KeyValuePair<string, object>("name", name),
                new KeyValuePair<string, object>("status", status),
                new KeyValuePair<string, object>("dateFrom", dateFrom),
                new KeyValuePair<string, object>("dateTo", dateTo),
                new KeyValuePair<string, object>("taskDefId", taskDefId),
                new KeyValuePair<string, object>("timeFrom", timeFrom),
                new KeyValuePair<string, object>("timeTo", timeTo)
            );
        }

        private void InsertTaskSchedulerToSites(string site, int taskSchId)
        {
            string query = @"
            DECLARE @siteId INT;
            SELECT TOP 1 @siteId = Id FROM sites WHERE Name LIKE @site;
 
            Insert into [dbo].[TaskSchedulersToSites]
            ([TaskSchedulerId]
              ,[SiteId])
	          values
	          (@taskSchId, @siteId);
            "
            ;
            ExecuteNonQuery(
            query,
                new KeyValuePair<string, object>("site", site),
                new KeyValuePair<string, object>("taskSchId", taskSchId)
            );
        }

        private int? TaskSchedulerToSiteExist(int taskSchId)
        {
            string query = @"
             select TaskSchedulerId from TaskSchedulersToSites where TaskSchedulerId = @taskSchId

              SELECT SCOPE_IDENTITY();
            "
            ;
            int? result = ExecuteAndGetInt(
             query,
                new KeyValuePair<string, object>("taskSchId", taskSchId)
            );
            return result == 0 ? (int?)null : result;
        }

        private void DeleteTaskScheduler(int taskSchId)
        {
            string query = @"
            delete from TaskSchedulers where Id = @taskSchId
            "
            ;
            ExecuteNonQuery(
            query,
                new KeyValuePair<string, object>("taskSchId", taskSchId)
            );
        }
    }
}
