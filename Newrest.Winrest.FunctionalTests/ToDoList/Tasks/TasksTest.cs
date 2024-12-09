using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Tasks;
using System;
using System.Collections.Generic;
using System.IO;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig;
using System.Linq;
using Newrest.Winrest.FunctionalTests.Utils;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Scheduler;
using System.Threading;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace Newrest.Winrest.FunctionalTests.ToDoList
{
    [TestClass]
    public class TasksTest : TestBase
    {
        private static Random random = new Random();
        private const int Timeout = 600000;
        private readonly string TASK_NAME = "Task_" + DateUtils.Now.ToString("dd/MM/yyyy") + "_" + random.Next(100, 9999);
   
        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(TD_TA_Filter_Search):
                    TestInitialize_CreateTask();
                    break;

                case nameof(TD_TA_Filter_Customers):
                    TestInitialize_CreateTask();
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
                case nameof(TD_TA_Filter_Search):
                    TestCleanup_DeleteTask();
                    break;

                case nameof(TD_TA_Filter_Customers):
                    TestCleanup_DeleteTask();
                    break;
                case nameof(TD_TA_Filter_ShowOnlyActive):
                    TestCleanup_DeleteTasksq();
                    break;
                default:
                    break;
            }
        }



        public void TestCleanup_DeleteTasksq()
        {
            
            Assert.IsNull(1, "La Task n'est pas supprimé.");
        }

        public void TestCleanup_DeleteTask()
        {
            DeleteTaskDefinition((int)TestContext.Properties[string.Format("TaskDefinitionId")]);

            Assert.IsNull(TaskDefinitionToSiteExist((int)TestContext.Properties[string.Format("TaskDefinitionId")]), "La Task n'est pas supprimé.");
        }


        public void TestInitialize_CreateTask()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            TestContext.Properties[string.Format("TaskDefinitionId")] = InsertTaskDefinition(customer, TASK_NAME);
            InsertTaskDefinitionToSites(site, (int)TestContext.Properties[string.Format("TaskDefinitionId")]);

            Assert.IsNotNull(TaskDefinitionToSiteExist((int)TestContext.Properties[string.Format("TaskDefinitionId")]), "La Task n'est pas crée.");
        }

        [Priority(0)]
        [TestMethod]
        [Timeout(Timeout)]
        [TestCategory("Prerequis")]
        public void TD_TA_CreateData()
        {
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            tasksPage.Filter(TasksPage.FilterType.Search, TASK_NAME);
            if (tasksPage.CheckTotalNumber() == 0)
            {
                tasksPage.CreateNewTaskWithMultipleSite(TASK_NAME, new List<string> { siteACE, siteMAD }, customer);
                tasksPage.BackToList();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, TASK_NAME);
            }
            Assert.IsTrue(tasksPage.CheckTotalNumber() > 0, "La nouveau task n'est pas ajoutée");

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_CreateNewTasks()
        {
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = "ACE";
            var customer = "3D COMM";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            var numberOfTasksInHeader = tasksPage.GetNumberOfTasksInHeader();
            tasksPage.CreateNewTask(name, site, customer);
            tasksPage.BackToList();
            tasksPage.ResetFilters();
            var numberOfTasksInHeaderAfterCreate = tasksPage.GetNumberOfTasksInHeader();
            Assert.AreEqual(int.Parse(numberOfTasksInHeaderAfterCreate), int.Parse(numberOfTasksInHeader) + 1, "Task n'est pas ajoutée");
            tasksPage.Filter(TasksPage.FilterType.Search, name);
            tasksPage.DeleteTask();
            tasksPage.ResetFilters();
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_DeleteTasksIndex()
        {
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = "ACE";
            var customer = "3D COMM";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();

            tasksPage.CreateNewTask(name, site, customer);
            tasksPage.BackToList();

            tasksPage.ResetFilters();
            tasksPage.Filter(TasksPage.FilterType.Search, name);

            var numberOfTasksInHeaderAfterCreate = tasksPage.CheckTotalNumber();
            tasksPage.DeleteTask();
            var numberOfTasksInHeader = tasksPage.CheckTotalNumber();
            Assert.AreEqual(numberOfTasksInHeader, numberOfTasksInHeaderAfterCreate - 1, "Task n'est pas supprimée");
            tasksPage.ResetFilters();
        }
        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_CreateNewStep()
        {
            //prepare
            string nameStep = "step test";
            string instruction = "instruction test";
            string comment = "comment test";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            //act
            tasksPage.ResetFilters();
            TasksGeneralInfo tasksGeneralInfo = tasksPage.SelectFirstTask();
            try
            {
                FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
                Assert.IsTrue(fiUpload.Exists, "Fichier d'entrée non trouve");
                tasksGeneralInfo.AddNewStep(nameStep, false, instruction, comment, fiUpload);
                Assert.IsTrue(tasksGeneralInfo.IsExistTaskDetails(), "Step n'est pas ajouté");
            }
            finally
            {
                tasksGeneralInfo.DeleteStep_TaskDetails();
                tasksGeneralInfo.BackToList();
            }
        }
        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_DeleteTasksDetail()
        {
            //prepare
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = "ACE";
            var customer = "3D COMM";
            string nameStep = "step test";
            string instruction = "instruction test";
            string comment = "comment test";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            try
            {
                TasksGeneralInfo generalInfo = tasksPage.CreateNewTask(name, site, customer);
                FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
                Assert.IsTrue(fiUpload.Exists, "Fichier d'entrée non trouve");
                generalInfo.AddNewStep(nameStep, false, instruction, comment, fiUpload);
                Assert.IsTrue(generalInfo.IsExistTaskDetails(), "Step n'est pas ajouté");
                generalInfo.DeleteStep_TaskDetails();
                Assert.IsTrue(generalInfo.VerifyDeleteTasksDetails(), "Tasks Details ne sont pas supprimées.");

            }
            finally
            {
                tasksPage.BackToList();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, name);
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_CreateFromOtherTasks()
        {
            Random r = new Random();
            var newDuplicateTaskName = "Duplicate task" + DateUtils.Now.ToString("dd/MM/yyyy") + r.Next(); ;
            string taskName = "Task_" + DateUtils.Now.ToString("dd/MM/yyyy") + r.Next(); ;
            //Arrange
            var homePage = LogInAsAdmin();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            var numberOfTasksInHeaderAfterCreate = tasksPage.GetNumberOfTasksInHeader();
            tasksPage.CreateFromOtherTasks(taskName, newDuplicateTaskName, false);
            tasksPage.BackToList();
            tasksPage.ResetFilters();
            var numberOfTasksInHeader = tasksPage.GetNumberOfTasksInHeader();
            Assert.AreEqual(int.Parse(numberOfTasksInHeader), int.Parse(numberOfTasksInHeaderAfterCreate) + 1, "Task n'est pas dupliquée");
            tasksPage.Filter(TasksPage.FilterType.Search, newDuplicateTaskName);
            tasksPage.DeleteTask();
            tasksPage.ResetFilters();
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_Filter_Search()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            TasksPage tasksPage = homePage.GoToDoList_Tasks();
           
            tasksPage.Filter(TasksPage.FilterType.Search, TASK_NAME);
            Assert.IsTrue(tasksPage.VerifySearchFilter(TASK_NAME), "Search filter ne fonctionne pas.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_Filter_Customers()
        {
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            tasksPage.Filter(TasksPage.FilterType.Customers, customer);
            bool isVerified = tasksPage.VerifyCustomerFilter(customer);    
            //Assert
            Assert.IsTrue(isVerified, "Customer filter ne fonctionne pas.");

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_Filter_ShowAll()
        {
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = TestContext.Properties["SiteACE"].ToString();
            var customer = TestContext.Properties["CustomerJob"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin(); ;

            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            if (tasksPage.CheckTotalNumber() < 20)
            {
                // Create
                tasksPage.CreateNewTask(name, site, customer);
                tasksPage.BackToList();
            }

            if (!tasksPage.isPageSizeEqualsTo100())

            {
                tasksPage.PageSize("8");
                tasksPage.PageSize("100");
            }

            tasksPage.Filter(TasksPage.FilterType.ShowOnlyActive, true);
            var nbActive = tasksPage.GetNumberOfTasksInHeader();
            tasksPage.Filter(TasksPage.FilterType.ShowOnlyInactive, true);
            var nbInactive = tasksPage.GetNumberOfTasksInHeader();
            tasksPage.Filter(TasksPage.FilterType.ShowAll, true);
            var nbTotal = tasksPage.GetNumberOfTasksInHeader();

            Assert.AreEqual(int.Parse(nbTotal), (int.Parse(nbActive) + int.Parse(nbInactive)), String.Format(MessageErreur.FILTRE_ERRONE, "'Show all'"));

            tasksPage.DeleteTask();
            tasksPage.ResetFilters();
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_Filter_ShowOnlyActive()
        {
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = "ACE";
            var customer = "3D COMM";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();

            if (tasksPage.CheckTotalNumber() < 20)
            {
                // Create
                tasksPage.CreateNewTask(name, site, customer);
                tasksPage.BackToList();
            }

            if (!tasksPage.isPageSizeEqualsTo100())
            {
                tasksPage.PageSize("8");
                tasksPage.PageSize("100");
            }

            tasksPage.Filter(TasksPage.FilterType.ShowOnlyActive, true);
            var nbActive = tasksPage.GetNumberOfTasksInHeader();
            tasksPage.Filter(TasksPage.FilterType.ShowOnlyInactive, true);
            var nbInactive = tasksPage.GetNumberOfTasksInHeader();
            tasksPage.Filter(TasksPage.FilterType.ShowAll, true);
            var nbTotal = tasksPage.GetNumberOfTasksInHeader();

            Assert.AreEqual(int.Parse(nbActive), (int.Parse(nbTotal) - int.Parse(nbInactive)), String.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));

            tasksPage.DeleteTask();
            tasksPage.ResetFilters();
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_Filter_ShowOnlyInactive()
        {
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = TestContext.Properties["SiteACE"].ToString();
            var customer = TestContext.Properties["CustomerJob"].ToString();

            //arrange
            HomePage homePage = LogInAsAdmin();

            //act
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();

            if (tasksPage.CheckTotalNumber() < 20)
            {
                // Create
                tasksPage.CreateNewTask(name, site, customer);
                tasksPage.BackToList();
            }

            if (!tasksPage.isPageSizeEqualsTo100())
            {
                tasksPage.PageSize("8");
                tasksPage.PageSize("100");
            }

            tasksPage.Filter(TasksPage.FilterType.ShowOnlyActive, true);
            var nbActive = tasksPage.GetNumberOfTasksInHeader();
            tasksPage.Filter(TasksPage.FilterType.ShowOnlyInactive, true);
            var nbInactive = tasksPage.GetNumberOfTasksInHeader();
            tasksPage.Filter(TasksPage.FilterType.ShowAll, true);
            var nbTotal = tasksPage.GetNumberOfTasksInHeader();

            Assert.AreEqual(int.Parse(nbInactive), (int.Parse(nbTotal) - int.Parse(nbActive)), String.Format(MessageErreur.FILTRE_ERRONE, "'Show only inActive'"));

            tasksPage.DeleteTask();
            tasksPage.ResetFilters();
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_Filter_Sites()
        {
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            //FIXME pas de ResetFiltrer, donc Search ici
            tasksPage.Filter(TasksPage.FilterType.Search, TASK_NAME);
            tasksPage.Filter(TasksPage.FilterType.Sites, siteACE);
            Assert.IsTrue(tasksPage.VerifySiteFilter(siteACE), "Site filter ne fonctionne pas.");

        }

        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_EditTasks()
        {
            //prepare
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = "ACE";
            var customer = "3D COMM";
            string nameStep = "step test";
            string instruction = "instruction test";
            string comment = "comment test";
            string nameStepEdited = "stepEdited";
            string instructionEdited = "instructionEdited";
            string commentEdited = "comment";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //act
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            TasksGeneralInfo generalInfo = tasksPage.CreateNewTask(name, site, customer);
            try
            {
                FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
                Assert.IsTrue(fiUpload.Exists, "Fichier d'entrée non trouve");
                generalInfo.AddNewStep(nameStep, false, instruction, comment, fiUpload);
                Assert.IsTrue(generalInfo.IsExistTaskDetails(), "Step n'est pas ajouté");
                generalInfo.EditTask(nameStepEdited, instructionEdited, commentEdited);
                Assert.IsTrue(generalInfo.VerifyEditTask(nameStepEdited, instructionEdited, commentEdited), "Edit task ne fonctionne pas.");
            }
            finally
            {
                tasksPage.BackToList();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, name);
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();

            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_ResetFilter()
        {
            object value = null;

            //Arrange
            HomePage homePage = LogInAsAdmin();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            int numberDefaultListSiteFilter = tasksPage.GetNumberSelectedSiteFilter();
            int numberDefaultListCustomerFilter = tasksPage.GetNumberSelectedCustomerFilter();
            //Filter
            tasksPage.Filter(TasksPage.FilterType.Search, "Test");
            tasksPage.Filter(TasksPage.FilterType.ShowAll, true);
            tasksPage.Filter(TasksPage.FilterType.ShowOnlyActive, true);
            tasksPage.Filter(TasksPage.FilterType.ShowOnlyInactive, true);
            tasksPage.Filter(TasksPage.FilterType.Sites, "ACE");
            tasksPage.Filter(TasksPage.FilterType.Customers, "3D COMM");
            //Reset filter
            tasksPage.ResetFilters();
            //Verify reset filter
            value = tasksPage.GetFilterValue(TasksPage.FilterType.Search);
            Assert.AreEqual("", value, "ResetFilter Search ''");
            value = tasksPage.GetFilterValue(TasksPage.FilterType.ShowAll);
            Assert.AreEqual(false, value, "ResetFilter ShowAll");
            value = tasksPage.GetFilterValue(TasksPage.FilterType.ShowOnlyActive);
            Assert.AreEqual(false, value, "ResetFilter ShowOnlyActive");
            value = tasksPage.GetFilterValue(TasksPage.FilterType.ShowOnlyInactive);
            Assert.AreEqual(true, value, "ResetFilter ShowOnlyInactive");
            var numberSelectedSiteFilter = tasksPage.GetNumberSelectedSiteFilter();
            Assert.AreEqual(numberSelectedSiteFilter, numberDefaultListSiteFilter, "ResetFilter Site ''");
            var numberSelectedCustomerFilter = tasksPage.GetNumberSelectedCustomerFilter();
            Assert.AreEqual(numberSelectedCustomerFilter, numberDefaultListCustomerFilter, "ResetFilter Customer ''");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_PrintQRCode()
        {
            //prepare
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = "ACE";
            var customer = "3D COMM";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_"; // "Test QR-Code";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            tasksPage.ClearDownloads();
            tasksPage.CreateNewTask(name, site, customer);
            try
            {
                tasksPage.PrintQRCode();
                //Open pdf file
                PrintReportPage reportPage = tasksPage.PrintItemGenerique(tasksPage);
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                Assert.IsTrue(fi.Exists, trouve + " non généré");
                //Vérifier le QR Code Exist sur le PDF
                PdfDocument document = PdfDocument.Open(fi.FullName);
                Page page1 = document.GetPage(1);
                List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
                Assert.AreEqual(1, images.Count, "QR Code n'existe pas.");
                IPdfImage image = images[0];
                var heightImg = image.HeightInSamples;
                var widthImg = image.WidthInSamples;
                Assert.AreEqual(heightImg, widthImg, 90, "QR Code size est faux.");
                //Verify name task
                IEnumerable<Word> wordPdf = page1.GetWords().ToList();
                string namePdf = wordPdf.FirstOrDefault().Text;
                Assert.AreEqual(namePdf, name, "Name Task est faux.");
            }
            finally
            {
                //Delete task
                homePage.GoToDoList_Tasks();
                tasksPage.ClearDownloads();
                tasksPage.Filter(TasksPage.FilterType.Search, name);
                tasksPage.Filter(TasksPage.FilterType.ShowOnlyActive, true);
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();
            }

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_ModifyTasks()
        {
            //Prepare
            Random r = new Random();
            string name = $"Task_{DateUtils.Now:yyyy-MM-dd}_{r.Next()}";
            string site = "ACE";
            string customer = "AIR CPU";
            //Arrange
            HomePage homePage = LogInAsAdmin();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            try
            {
                //Act
                TasksGeneralInfo generalInfo = tasksPage.CreateNewTask(name, site, customer);
                //Etre sur une Task
                //"1. Onglet Général info
                //2.Modifier tous les champs
                generalInfo.SetName($"{name}_renamed");
                generalInfo.SetSite("MAD");
                generalInfo.SetCustomer("AIR CANADA");
                generalInfo.SetVehiculeRegistration("Nothing Selected");
                generalInfo.SetEnforceOrder(true);
                generalInfo.SetAirCraft("B777");
                generalInfo.SetDescription("hello description");
                generalInfo.SetComment("hello comment");
                generalInfo.SetActive(false);
                generalInfo.SetDuration("30");
                //3.Vérifer que les modifications soient bien prise en compte"
                generalInfo.ClickOnSchedulersTab();
                generalInfo.ClickOnGeneralInfoTab();
                Assert.AreEqual(generalInfo.GetName(), $"{name}_renamed");
                Assert.IsTrue(generalInfo.GetSite("MAD"));
                Assert.AreEqual(generalInfo.GetCustomer(), "AIR CANADA");
                Assert.AreEqual(generalInfo.GetVehiculeRegistration(), "Nothing Selected");
                Assert.AreEqual(generalInfo.IsEnforceOrder(), true);
                Assert.IsTrue(generalInfo.GetAirCraft("B777"));
                Assert.AreEqual(generalInfo.GetDescription(), "hello description");
                Assert.AreEqual(generalInfo.GetComment(), "hello comment");
                Assert.AreEqual(generalInfo.IsActive(), false);
                Assert.AreEqual(generalInfo.GetDuration(), "30");
                tasksPage = generalInfo.BackToList();
            }
            finally
            {
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.ShowAll, true);
                tasksPage.Filter(TasksPage.FilterType.Search, $"{name}_renamed");
                tasksPage.DeleteTask();
            }
        }
        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_FoldAndUnfold()
        {
            //Prepare
            Random r = new Random();
            string name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            string site = "ACE";
            string customer = "AIR CPU";
            string nameStep = "step test";
            string instruction = "instruction test";
            string comment = "comment test";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            //Act
            TasksGeneralInfo generalInfo = tasksPage.CreateNewTask(name, site, customer);
            try
            {
                Console.WriteLine(TestContext.TestDeploymentDir + "\\pizza.png");
                FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
                Assert.IsTrue(fiUpload.Exists, "Fichier d'entrée non trouve");
                generalInfo.AddNewStep(nameStep, false, instruction, comment, fiUpload);
                Assert.IsTrue(generalInfo.IsExistTaskDetails(), "Step n'est pas ajouté");
                generalInfo.BackToList();
                tasksPage.Filter(TasksPage.FilterType.Search, name);
                tasksPage.FoldAndUnfold();
                Assert.IsTrue(tasksPage.VerifyFoldAndUnfold(), "Fold and Unfold ne fonctionne pas.");
                tasksPage.FoldAndUnfold();
                Assert.IsFalse(tasksPage.VerifyFoldAndUnfold(), "Fold and Unfold ne fonctionne pas.");

            }
            finally
            {
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();
            }
        }

        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_CreateNewStepWithPicture()
        {
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = "ACE";
            var customer = "3D COMM";
            string nameStep = "step test";
            string instruction = "instruction test";
            string comment = "comment test";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            try
            {
                TasksGeneralInfo generalInfo = tasksPage.CreateNewTask(name, site, customer);
                generalInfo.BackToList();
                tasksPage.Filter(TasksPage.FilterType.Search, name);
                tasksPage.SelectFirstTask();
                Console.WriteLine(TestContext.TestDeploymentDir + "\\pizza.png");
                FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
                Assert.IsTrue(fiUpload.Exists, "Fichier d'entrée non trouve");
                generalInfo.AddNewStep(nameStep, false, instruction, comment, fiUpload);
                Assert.IsTrue(generalInfo.IsExistTaskDetails(), "Step n'est pas ajouté");
            }
            finally
            {
                tasksPage.BackToList();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, name);
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_DuplicateTask()
        {
            //prepare
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            string site = TestContext.Properties["SiteACE"].ToString();
            string customer = "3D COMM";
            var newTaskName = name + "_" + "COPY";
            bool duplicateSchedule = true;
            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //act
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            TasksGeneralInfo generalInfo = tasksPage.CreateNewTask(name, site, customer);
            try
            {
                tasksPage = generalInfo.BackToList();
                tasksPage.ResetFilters();
                tasksPage.CreateFromOtherTasks(name, newTaskName, duplicateSchedule);
                Assert.AreEqual(generalInfo.GetSiteValue(), site, "Duplicate ne fonctionne pas, site non compatible.");
                Assert.AreEqual(generalInfo.GetCustomer(), customer, "Duplicate ne fonctionne pas, customer non compatible.");
                Assert.AreEqual(generalInfo.GetName(), newTaskName, "Duplicate ne fonctionne pas.");
            }
            finally
            {
                tasksPage = generalInfo.BackToList();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, newTaskName);
                tasksPage.DeleteTask();
                tasksPage.Filter(TasksPage.FilterType.Search, name);
                tasksPage.DeleteTask();
                tasksPage.ResetFilters();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_ScanOnly()
        {
            Random r = new Random();
            var name = "Task_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + r.Next();
            var site = TestContext.Properties["SiteACE"].ToString();
            var customer = TestContext.Properties["CustomerJob"].ToString(); ;
            bool value = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            try
            {
                tasksPage.ResetFilters();
                var numberOfTasksInHeader = tasksPage.GetNumberOfTasksInHeader();
                TasksGeneralInfo tasksGeneralInfo = tasksPage.CreateNewTask(name, site, customer, false, false, true);
                tasksPage.BackToList();
                tasksPage.Filter(TasksPage.FilterType.Search, name);
                numberOfTasksInHeader = tasksPage.GetNumberOfTasksInHeader();
                Assert.AreEqual(int.Parse(numberOfTasksInHeader), 0, "Le task a été crée.");
            }
            finally
            {
                homePage.GoToDoList_Tasks();
                tasksPage.Filter(TasksPage.FilterType.Search, name);
                if (tasksPage.CheckTotalNumber() != 0)
                {
                    tasksPage.DeleteTask();
                    tasksPage.ResetFilters();
                }
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TD_TA_DuplicateTaskAndScheduler()
        {
            //prepare
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string taskName = TASK_NAME + "- For Test DUPLICATE";
            string schelduerName = "Schedule" + "- For Test DUPLICATE";
            var statusProduction = TestContext.Properties["StatusProduction"].ToString();
            var newTaskName = taskName + "_" + "COPY";
            bool duplicateSchedule = true;
            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            TasksPage tasksPage = homePage.GoToDoList_Tasks();
            tasksPage.ResetFilters();
            TasksGeneralInfo tasksGeneralInfo = tasksPage.CreateNewTaskWithMultipleSite(taskName, new List<string> { siteACE, siteMAD }, customer);
            try
            {
                tasksGeneralInfo.ClickOnSchedulersTab();
                tasksGeneralInfo.ShowPlusScheduler();
                tasksGeneralInfo.CreateSchedulerModalPage();
                tasksGeneralInfo.CreateNewSchedule(schelduerName, taskName, siteACE, DateUtils.Now, statusProduction,
                        DateUtils.Now.AddDays(-1), DateUtils.Now.AddDays(1), "2");

                tasksPage.BackToList();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.CreateFromOtherTasks(taskName, newTaskName, duplicateSchedule);
                tasksPage.BackToList();

                SchedulerPage schedulerPage = homePage.GoToDoList_Scheduler();
                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchTaskName, taskName);
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                Assert.AreNotEqual(0, "SCHEDULER n'est pas Compatible au filtrage");

                schedulerPage.ResetFilters();
                schedulerPage.Filter(SchedulerPage.FilterType.SearchTaskName, newTaskName);
                schedulerPage.Filter(SchedulerPage.FilterType.SearchSchedulerName, schelduerName);
                Assert.AreNotEqual(0, "SCHEDULER DUPLICATE n'est pas Compatible au filtrage");
            }
            finally
            {
                tasksPage = homePage.GoToDoList_Tasks();
                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, taskName);
                tasksPage.DeleteTask();

                tasksPage.ResetFilters();
                tasksPage.Filter(TasksPage.FilterType.Search, newTaskName);
                tasksPage.DeleteTask();
            }
        }


        ////////////////////////////////Insert Task////////////////////////////////

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

    }
}