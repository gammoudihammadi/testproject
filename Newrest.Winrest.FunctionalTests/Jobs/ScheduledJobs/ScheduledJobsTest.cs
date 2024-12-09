using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualBasic.Logging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.Jobs
{
    [TestClass]
    public class ScheduledJobsTest : TestBase
    {
        private const int _timeout = 600000;
        private readonly string CUSTOMER = "3D COMM";
        private readonly string CONVERTER = "AC_HR_txt";
        private readonly string TYPE = "fileflowtype";

        string FILE_FLOW_NAME = "FileFlow" + DateUtils.Now.ToString("dd/MM/yyyy");

        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void JB_SCHED_index_CreateData()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobscheduled.Nav_To_File_Flows();
            scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowAll, true);

            //if (scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage() != 0)
            //{
            //    scheduledJobsTabFileFlowsPage.DeleteFirstFileFlow();
            //}

            var fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
            fileFlowsCreateModalPage.FillFiled_CreateFileFlows(FILE_FLOW_NAME, CUSTOMER, true, null, CONVERTER);
            scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, FILE_FLOW_NAME);
            scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.Customers, CUSTOMER);
            scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.Converter, CONVERTER);
            Assert.IsTrue(scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage() >= 1, "Le File Flows n'est pas créé.");
        }

        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void JB_SCHED_Cegid_CreateData()
        {
            //Prepare
            string name = "CegidNewRest2024";
            string fiscalEntities = "NEWREST ESPAÑA S.L.";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobsPage.Nav_To_Cegid();
            scheduledJobsCegidPage.ResetFilter();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, name);
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowActiveOnly, true);
            if (!scheduledJobsCegidPage.GetAllNamesResultPaged().Contains(name))
            {
                CegidScheduledJobModal newCegidScheduledModal = scheduledJobsCegidPage.NewCegidScheduledModal();
                newCegidScheduledModal.FillFieldsModalCegidScheduledJob(name, fiscalEntities);
            }
            scheduledJobsCegidPage.ResetFilter();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, name);
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowActiveOnly, true);

            //Assert
            bool cegitCreated = scheduledJobsCegidPage.GetAllNamesResultPaged().Contains(name);
            Assert.IsTrue(cegitCreated, "Le Cegid n'a pas été créé.");

        }
        [Priority(2)]
        [TestMethod]
        [Timeout(_timeout)]
        public void JB_SCHED_Bank_CreateData()
        {
            string nameBank = "BankNewRest2024";
            string companyCode = "testCompanyCod";
            string sftpConnectionString = "testSFTPstring";
            string sftpFolder = "testSFTPFolderstring";
            string customer = CUSTOMER; //"1Customer for Production";
            string converterType = "Default";

            // Arrange
            var homePage = LogInAsAdmin();

            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ResetFilters();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.SearchByName, nameBank);
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowActiveOnly, true);
            if (!scheduledJobsBankPage.GetAllNamesPageResult().Contains(nameBank))
            {
                scheduledJobsBankPage.AddNewBank(nameBank, companyCode, sftpConnectionString, sftpFolder, true, customer, converterType);
            }
            scheduledJobsBankPage.ResetFilters();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.SearchByName, nameBank);
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowActiveOnly, true);

            //Assert
            bool nameBankExists = scheduledJobsBankPage.GetAllNamesPageResult().Contains(nameBank);
            Assert.IsTrue(nameBankExists, "Le Bank n'a pas été créé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_Delete()
        {
            //Prepare
            string namefordelete = "NameForShowValidator" + new Random().Next().ToString();
            string customer = "3D COMM";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobscheduled.Nav_To_File_Flows();

            try
            {
                //create
                var fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
                FileFlowsCreateModalPage fileFlowsPage = fileFlowsCreateModalPage.FillFiled_CreateFileFlows(namefordelete, customer, true);
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, namefordelete);
                scheduledJobsTabFileFlowsPage.ClickSurPoubelleDeleteConfirm(namefordelete);
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, namefordelete);
                var listElemnetsBeforeAfter = scheduledJobsTabFileFlowsPage.GetElementsNamesFileFlowProviders();

                //Assert
                bool nameExist = listElemnetsBeforeAfter.Contains(namefordelete);
                Assert.IsFalse(nameExist, "Delete Was Failed");
            }

            finally
            {
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, namefordelete);
                var listElemnetsBeforeAfter = scheduledJobsTabFileFlowsPage.GetElementsNamesFileFlowProviders();
                if (listElemnetsBeforeAfter.Contains(namefordelete))
                {
                    scheduledJobsTabFileFlowsPage.ClickSurPoubelleDeleteConfirm(namefordelete);
                }

            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_DeleteValidator()
        {
            //Prepare
            string namefordelete = "NameForShowValidator" + new Random().Next().ToString();
            string customer = "3D COMM";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobscheduled.Nav_To_File_Flows();
            try
            {
                var fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
                FileFlowsCreateModalPage fileFlowsPage = fileFlowsCreateModalPage.FillFiled_CreateFileFlows(namefordelete, customer, true);
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, namefordelete);
                scheduledJobsTabFileFlowsPage.ClickSurPoubelle();
                bool popupVisible = scheduledJobsTabFileFlowsPage.CheckPopupFileFlowsDeleteIsVisible();

                //Assert
                Assert.IsTrue(popupVisible, "La pop-up pour demander confirmation ou annulation de l'action n'apparaît pas");
            }
            finally
            {
                scheduledJobsTabFileFlowsPage.DeleteCancel();
                scheduledJobsTabFileFlowsPage.DeleteFirstFileFlow();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Bank_filter_ResetFilter()
        {
            //Prepare
            string nameBank = "testName" + DateTime.Now.ToString("dd/MM/yyyy");
            string companyCode = "testCompanyCod";
            string sftpConnectionString = "testSFTPstring";
            string sftpFolder = "testSFTPFolderstring";
            string customer = TestContext.Properties["CustomerDelivery"].ToString(); 
            string converterType = "Default";

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ResetFilters();
            if (scheduledJobsBankPage.GetTotalResultRowsPage() == 0 || (scheduledJobsBankPage.GetAllPageResultConverterType().All(item => item != converterType)))
                scheduledJobsBankPage.AddNewBank(nameBank, companyCode, sftpConnectionString, sftpFolder, true, customer, converterType);
            var firstLineToDeleteName = scheduledJobsBankPage.GetFirstLineName();
            scheduledJobsBankPage.ResetFilters();
            var defaultShow = scheduledJobsBankPage.GetShowFilterSelected();
            var defaultCustomersSelected = scheduledJobsBankPage.GetSelectedCustomersToFilter();
            var defaultConverterTypesSelected = scheduledJobsBankPage.GetConverterTypesSelectedToFilter();
            var defaultItemsDisplayed = scheduledJobsBankPage.GetNumberItemsDisplayed();

            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.SearchByName, nameBank);
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ConverterTypes, converterType);
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.Customer, customer);
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowAll, true);
            scheduledJobsBankPage.ResetFilters();

            var resultShowAfterReset = scheduledJobsBankPage.GetShowFilterSelected();
            var resultCustomersSelectedAfterReset = scheduledJobsBankPage.GetSelectedCustomersToFilter();
            var resultConverterTypesSelectedAfterReset = scheduledJobsBankPage.GetConverterTypesSelectedToFilter();
            var resultItemsDisplayedAfterResult = scheduledJobsBankPage.GetNumberItemsDisplayed();

            //Assert
            Assert.AreEqual(defaultShow, resultShowAfterReset, "Le filter Show ne remet pas");
            Assert.AreEqual(defaultCustomersSelected.Count, resultCustomersSelectedAfterReset.Count, "Le filter CUSTOMERS ne remet pas");
            Assert.AreEqual(defaultConverterTypesSelected.Count, resultConverterTypesSelectedAfterReset.Count, "Le filter CONVERTER TYPES ne remet pas");
            Assert.AreEqual(defaultItemsDisplayed, resultItemsDisplayedAfterResult, "Le filter  ne remet pas");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_bank_AddNew()
        {
            string rnd = new Random().Next(10, 99).ToString();
            var customer = TestContext.Properties["CustomerJob"].ToString();
            var name = "testName" + rnd + DateTime.Now.ToString("dd/MM/yyyy");
            var companyCode = "testCompanyCode";
            var sftConnectionString = "testSFTPstring";
            var sftpFolder = "testSFTPstring";
            bool isActive = true;
            string converterType = "Default";
            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            try
            {
                scheduledJobsBankPage.AddNewBank(name, companyCode, sftConnectionString, sftpFolder, isActive, customer, converterType);
                scheduledJobsBankPage.ResetFilters();
                scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.SearchByName, name);
                var verifyName = scheduledJobsBankPage.GetFirstLineName().Equals(name);
                var verifyCustomer = scheduledJobsBankPage.GetFirstLineCustomer().Equals(customer);
                var verifyCompanyCode = scheduledJobsBankPage.GetFirstLineCompanyCode().Equals(companyCode);
                var verifyConverterType = scheduledJobsBankPage.GetFirstLineConverterType().Equals(converterType);
                var dataCorrect = verifyName && verifyCustomer && verifyCompanyCode && verifyConverterType;
                Assert.IsTrue(dataCorrect, "L'ajout d'un nouveau Bank ne fonctionne pas correctement!");
            }
            finally
            {
                scheduledJobsBankPage.DeleteFirstBank();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_AddNew_SpecificJob()
        {
            Random rnd = new Random();
            var email = TestContext.Properties["Admin_UserName"].ToString() + rnd.Next().ToString();
            
            var homePage = LogInAsAdmin();
            // Act
            SettingsPage jobsPage = homePage.GoToJobs_Settings();
            JobNotifications jobNotifications = jobsPage.GoTo_JOB_NOTIFICATION();
            try
            {
                CreateSpecificJobNotificationModel createSpecificJobNotificationModel = jobNotifications.AddNewSpecificJobNotification();
                jobNotifications = createSpecificJobNotificationModel.FillField_AddNewspecificjobnotification(email);
                jobNotifications.Filter(JobNotifications.FilterType.Search, email);
                jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, "Specific");
                Assert.IsTrue(jobNotifications.FindRowEmail(email), "La notification n'apparait pas dans la liste ");
            }
            finally
            {
                jobNotifications.DeleteFirstJob();
            }


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_AddNew_SpecificJobValidator()
        {
            
            var homePage = LogInAsAdmin();
            // Act
            SettingsPage jobsPage = homePage.GoToJobs_Settings();
            JobNotifications jobNotifications = jobsPage.GoTo_JOB_NOTIFICATION();
            CreateSpecificJobNotificationModel createSpecificJobNotificationModel = jobNotifications.AddNewSpecificJobNotification();
            createSpecificJobNotificationModel.FillField_AddNewspecificjobnotification("");
            var isEmailRequired = createSpecificJobNotificationModel.VerifyEmailRequired();
            Assert.IsTrue(isEmailRequired, "Les validators n'apparaissent pas");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_bank_AddNewValidator()
        {
            var homePage = LogInAsAdmin();
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.OpenModalAddNewBank();
            scheduledJobsBankPage.ClickCreate();
            var validatorsExist = scheduledJobsBankPage.VerifyCreateValidators();
            Assert.IsTrue(validatorsExist, "Les validators n'apparaissent pas dans le pop-up.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Bank_Delete()
        {
            //Prepare 
            string nameBank = "testName" + DateTime.Now.ToString();
            string companyCode = "testCompanyCod";
            string sftpConnectionString = "testSFTPstring";
            string sftpFolder = "testSFTPFolderstring";
            string customer = TestContext.Properties["CustomerDelivery"].ToString();
            string converterType = "Default";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ResetFilters();
            if (scheduledJobsBankPage.GetTotalResultRowsPage() == 0 || (scheduledJobsBankPage.GetAllPageResultConverterType().All(item => item != converterType)))
                scheduledJobsBankPage.AddNewBank(nameBank, companyCode, sftpConnectionString, sftpFolder, true, customer, converterType);
            var firstLineToDeleteName = scheduledJobsBankPage.GetFirstLineName();
            scheduledJobsBankPage.DeleteFirstBank();
            scheduledJobsBankPage.ResetFilters();
            if (scheduledJobsBankPage.GetTotalResultRowsPage() > 0)
            {
                var firstLineAfterDeletionName = scheduledJobsBankPage.GetFirstLineName();
                Assert.AreNotEqual(firstLineToDeleteName, firstLineAfterDeletionName, "La ligne n'a pas été supprimée correctement.");
            }
            else
            {
                Assert.IsTrue(true, "La ligne a été supprimée et aucune ligne restante.");
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Bank_Edit()
        {
            string rnd = new Random().Next(100, 999).ToString();
            string newName = "EditedName" + rnd;

            var homePage = LogInAsAdmin();
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ResetFilters();
            scheduledJobsBankPage.ClickFirstLine();
            scheduledJobsBankPage.SetName(newName);
            scheduledJobsBankPage.ClickCreate();
            var editedName = scheduledJobsBankPage.GetFirstLineName();
            Assert.AreEqual(editedName, newName, "Les informations ne sont pas modifiées");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Bank_filter_ConverterTypes()
        {
            //Prepare
            string nameBank = "testName" + DateTime.Now.ToString();
            string companyCode = "testCompanyCod";
            string sftpConnectionString = "testSFTPstring";
            string sftpFolder = "testSFTPFolderstring";
            string customer = TestContext.Properties["CustomerDelivery"].ToString();
            string converterType = "Default";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ResetFilters();
            if (scheduledJobsBankPage.GetTotalResultRowsPage() == 0 || (scheduledJobsBankPage.GetAllPageResultConverterType().All(item => item != converterType)))
                scheduledJobsBankPage.AddNewBank(nameBank, companyCode, sftpConnectionString, sftpFolder, true, customer, converterType);
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ConverterTypes, converterType);

            //Assert
            bool typeVerified = scheduledJobsBankPage.GetAllPageResultConverterType().All(item => item == converterType);
            Assert.IsTrue(typeVerified, "Les résultats ne sont pas mis à jour en fonction des filtres par Converter Types");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Bank_filter_Customers()
        {
            string nameBank = "Bank" + DateTime.Now.ToString();
            string companyCode = "testCompanyCod";
            string sftpConnectionString = "testSFTPstring";
            string sftpFolder = "testSFTPFolderstring";
            var customer = TestContext.Properties["CustomerJob"].ToString();
            string converterType = "Default";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ResetFilters();
            if (scheduledJobsBankPage.GetTotalResultRowsPage() == 0 || (scheduledJobsBankPage.GetAllPageResultCustomers().All(item => item != customer)))
                scheduledJobsBankPage.AddNewBank(nameBank, companyCode, sftpConnectionString, sftpFolder, true, customer, converterType);
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.Customer, customer);

            //Assert
            bool customerFilterVerified = scheduledJobsBankPage.GetAllPageResultCustomers().All(item => item == customer) || scheduledJobsBankPage.GetAllPageResultCustomers().Count == 0;
            Assert.IsTrue(customerFilterVerified, "Les résultats ne sont pas mis à jour en fonction des filtres par Customers");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Bank_filter_ShowAll()
        {

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ResetFilters();
            scheduledJobsBankPage.PageSize("100");
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowActiveOnly, true);
            var resultsActivated = scheduledJobsBankPage.GetTotalResultRowsPage();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowInactiveOnly, true);
            var resultsInactivated = scheduledJobsBankPage.GetTotalResultRowsPage();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowAll, true);
            var resultAll = scheduledJobsBankPage.GetTotalResultRowsPage();
            Assert.AreEqual(resultsActivated + resultsInactivated, resultAll, "Les résultats ne sont pas mis à jour en fonction des filtres Show All");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_filter_ShowInactiveOnly()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobsPage.Nav_To_File_Flows();
            scheduledJobsTabFileFlowsPage.ResetFilter();
            scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowInactiveOnly, true);

            //Assert
            bool allRowsCheckedInactive = scheduledJobsTabFileFlowsPage.CheckAllRowsIsActiveOrInactive(false);
            Assert.IsTrue(allRowsCheckedInactive, "Filter show inactive only n'est pas fonctionnel");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Bank_filter_ShowActiveOnly()
        {

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ResetFilters();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowActiveOnly, true);
            scheduledJobsBankPage.PageSize("100");
            var resultsActivated = scheduledJobsBankPage.GetTotalResultRowsPage();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowAll, true);
            var resultAll = scheduledJobsBankPage.GetTotalResultRowsPage();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowInactiveOnly, true);
            var resultsInactivated = scheduledJobsBankPage.GetTotalResultRowsPage();

            //Assert
            Assert.AreEqual(resultAll - resultsInactivated, resultsActivated, "Les résultats ne sont pas mis à jour en fonction des filtres Show ActiveOnly");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Bank_Pagination()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsPage = homePage.GoToJobs_ScheduledJobs();
            var scheduledJobsBankPage = jobsPage.Nav_To_Bank();

            scheduledJobsBankPage.PageSize("8");
            var count = scheduledJobsBankPage.GetAllNamesPageResult().Count;
            Assert.IsTrue(count <= 8, "Paggination ne fonctionne pas..");
            scheduledJobsBankPage.PageSize("16");
            count = scheduledJobsBankPage.GetAllNamesPageResult().Count;
            Assert.IsTrue(count <= 16, "Paggination ne fonctionne pas..");
            scheduledJobsBankPage.PageSize("30");
            count = scheduledJobsBankPage.GetAllNamesPageResult().Count;
            Assert.IsTrue(count <= 30, "Paggination ne fonctionne pas..");
            scheduledJobsBankPage.PageSize("50");
            count = scheduledJobsBankPage.GetAllNamesPageResult().Count;
            Assert.IsTrue(count <= 50, "Paggination ne fonctionne pas..");
            scheduledJobsBankPage.PageSize("100");
            count = scheduledJobsBankPage.GetAllNamesPageResult().Count;
            Assert.IsTrue(count <= 100, "Paggination ne fonctionne pas..");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Bank_filter_ShowInactiveOnly()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ResetFilters();
            scheduledJobsBankPage.PageSize("100");
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowActiveOnly, true);
            var resultsActivated = scheduledJobsBankPage.GetTotalResultRowsPage();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowAll, true);
            var resultAll = scheduledJobsBankPage.GetTotalResultRowsPage();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.ShowInactiveOnly, true);
            var resultsInactivated = scheduledJobsBankPage.GetTotalResultRowsPage();
            Assert.AreEqual(resultAll - resultsActivated, resultsInactivated, "Les résultats ne sont pas mis à jour en fonction des filtres Show InactiveOnly");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_filter_Customers()
        {
            //Prepare
            string fileFlow = "Runfile_test";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            jobsPage.Filter(ScheduledJobsPage.FilterType.ShowAll, true);
            try
            {
                ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobsPage.Nav_To_File_Flows();
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowAll, true);

                //create file
                var fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
                fileFlowsCreateModalPage.FillFiled_CreateFileFlows(fileFlow, CUSTOMER, true, null, CONVERTER);
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlow);
                int totalResultRowsPage = scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage();
                Assert.IsTrue(totalResultRowsPage >= 1, "Le File Flows n'est pas créé.");

                jobsPage.Filter(ScheduledJobsPage.FilterType.Customer, CUSTOMER);
                var verifyCustomerFilter = jobsPage.CheckAllRowsIsSpecificCustomer(CUSTOMER);

                //Assert
                Assert.IsTrue(verifyCustomerFilter, "Le Filter n'est pas fonctionnel ! ");
            }
            finally
            {
                //delete file
                jobsPage.DeleteFirstFileFlow();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_filter_ShowAll()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobscheduled.Nav_To_File_Flows();
            scheduledJobsTabFileFlowsPage.ResetFilter();
            scheduledJobsTabFileFlowsPage.PageSize("100");

            scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowActiveOnly, true);
            var resultsActivated = scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage();
            scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowInactiveOnly, true);
            var resultsInactivated = scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage();
            scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowAll, true);
            var resultAll = scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage();

            //Assert
            Assert.AreEqual(resultsActivated + resultsInactivated, resultAll, "Les résultats ne sont pas mis à jour en fonction des filtres Show All");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_filter_ShowActiveOnly()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobsPage.Nav_To_File_Flows();
            scheduledJobsTabFileFlowsPage.ResetFilter();
            scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowActiveOnly, true);

            var allRowsIsActive = scheduledJobsTabFileFlowsPage.CheckAllRowsIsActiveOrInactive(true);

            //assert
            Assert.IsTrue(allRowsIsActive, "Les résultats ne sont pas mis à jour en fonction des filtres Show Active Only");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_RunConfirm()
        {
            //Prepare
            string fileFlow = "Runfile_test";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage scheduledJobsPage = homePage.GoToJobs_ScheduledJobs();
            if (scheduledJobsPage.GetTotalResultRowsPage() >= 1)
            {
                scheduledJobsPage.RunJob();
                var isShowed = scheduledJobsPage.CheckOpenedConfirmRunJobModal();

                //Assert
                Assert.IsTrue(isShowed, "La pop-up n'a pas apparu!");
            }
            else
            {
                try
                {
                    ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = scheduledJobsPage.Nav_To_File_Flows();
                    scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowAll, true);

                    //create a FileFlow
                    var fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
                    fileFlowsCreateModalPage.FillFiled_CreateFileFlows(fileFlow, CUSTOMER, true, null, CONVERTER);
                    scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlow);
                    Assert.IsTrue(scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage() >= 1, "Le File Flows n'est pas créé.");
                    scheduledJobsPage.RunJob();
                    var isShowed = scheduledJobsPage.CheckOpenedConfirmRunJobModal();
                    Assert.IsTrue(isShowed, "La pop-up n'a pas apparu!");
                }
                finally
                {
                    // delete the FileFlow created
                    scheduledJobsPage.CancelForJob();
                    scheduledJobsPage.DeleteFirstFileFlow();
                }
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_filter_SearchByName()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobsPage.Nav_To_Cegid();
            var cegidName = scheduledJobsCegidPage.GetFirstCegidName();
            scheduledJobsCegidPage.ResetFilter();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, cegidName);

            //Assert
            bool cegidNameExist = scheduledJobsCegidPage.GetAllNamesResultPaged().Contains(cegidName);
            Assert.IsTrue(cegidNameExist, "Les résultats ne sont pas mis à jour en fonction des filtres.");

        }


        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_Run()
        {
            //Prepare
            string fileFlow = "Runfile_test";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage scheduledJobsPage = homePage.GoToJobs_ScheduledJobs();
           
            if (scheduledJobsPage.GetTotalResultRowsPage() >= 1)
            {
                scheduledJobsPage.RunJob();
                scheduledJobsPage.ConfirmRunJob();

                bool isShowed = scheduledJobsPage.CheckOpenedStatusJobPage();
                Assert.IsTrue(isShowed, "La page contenant le status de job ne s'ouvre pas!");
            }
            else
            {
                try
                {
                    ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = scheduledJobsPage.Nav_To_File_Flows();
                    scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowAll, true);

                    //create FileFLows
                    var fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
                    fileFlowsCreateModalPage.FillFiled_CreateFileFlows(fileFlow , CUSTOMER, true, null, CONVERTER);
                    scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlow);
                    Assert.IsTrue(scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage() >= 1, "Le File Flows n'est pas créé.");
                    scheduledJobsPage.RunJob();
                    scheduledJobsPage.ConfirmRunJob();

                    var isShowed = scheduledJobsPage.CheckOpenedStatusJobPage();
                    Assert.IsTrue(isShowed, "La page contenant le status de job ne s'ouvre pas!");
                }
                finally
                {
                    //delete the FileFLows created
                    scheduledJobsPage.Nav_To_Scheduled();
                    scheduledJobsPage.DeleteFirstFileFlow();
                }
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_FileFlows_CronExpressionEditor()
        {
            string fileFlow = "Runfile_test";
            
            var homePage = LogInAsAdmin();
            ScheduledJobsPage scheduledJobsPage = homePage.GoToJobs_ScheduledJobs();
            if (scheduledJobsPage.GetTotalResultRowsPage() >= 1)
            {
                scheduledJobsPage.Nav_To_File_Flows();
                scheduledJobsPage.ClickOnFileFlowItem();
                scheduledJobsPage.ClickOnEditorButton();
                var exist = scheduledJobsPage.CheckTabsExistance();
                Assert.IsTrue(exist, "Les onglets ne s'affichent pas!");
            }
            else
            {
                try
                {
                    ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = scheduledJobsPage.Nav_To_File_Flows();
                    scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowAll, true);

                    var fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
                    fileFlowsCreateModalPage.FillFiled_CreateFileFlows(fileFlow, CUSTOMER, true, null, CONVERTER);
                    scheduledJobsTabFileFlowsPage.ResetFilter();
                    scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlow);
                    var isAdded = scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage() >= 1;
                    Assert.IsTrue(isAdded, "Le File Flows n'est pas créé.");
                    scheduledJobsPage.ClickOnFileFlowItem();
                    scheduledJobsPage.ClickOnEditorButton();
                    var exist = scheduledJobsPage.CheckTabsExistance();
                    Assert.IsTrue(exist, "Les onglets ne s'affichent pas!");

                }
                finally
                {
                    scheduledJobsPage.CloseForEdit(); 
                    scheduledJobsPage.DeleteFirstFileFlow();
                }
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_filter_ResetFilter()
        {
            //Prepare
            string name = "CegidNewRest2024";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobsPage.Nav_To_Cegid();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, name);
            var countBeforeResetFilter = scheduledJobsCegidPage.GetAllNamesResultPaged().Count;
            scheduledJobsCegidPage.ResetFilter();
            var countAfterResetFilter = scheduledJobsCegidPage.GetAllNamesResultPaged().Count;

            //Assert
            Assert.AreNotEqual(countBeforeResetFilter, countAfterResetFilter, "Filter reset failed");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_filter_SearchByName()
        {
            //Prepare
            string nameBank = "FileFlow" + DateUtils.Now.ToString("dd/MM/yyyy");
            string companyCode = "testCompanyCod";
            string sftpConnectionString = "testSFTPstring";
            string sftpFolder = "testSFTPFolderstring";
            string customer = TestContext.Properties["CustomerDelivery"].ToString();
            string converterType = "Default";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ResetFilters();
            if (scheduledJobsBankPage.GetTotalResultRowsPage() != 0 || scheduledJobsBankPage.GetTotalResultRowsPage() == 0 || (scheduledJobsBankPage.GetAllPageResultConverterType().All(item => item != converterType)))
                scheduledJobsBankPage.AddNewBank(nameBank, companyCode, sftpConnectionString, sftpFolder, true, customer, converterType);
            scheduledJobsBankPage.ResetFilters();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.SearchByName, FILE_FLOW_NAME);

            //Assert
            int numberRow= scheduledJobsBankPage.NombreRow();
            Assert.IsTrue(numberRow >= 1, "Les lignes ayant un Name correspondant à la recherche n'apparaitre pas.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_AddNew()
        {
            string rnd = new Random().Next(1000, 5000).ToString();
            string name = "AddNewCegidForTest" + rnd;
            string fiscalEntites = "NEWREST ESPAÑA S.L.";

            var homePage = LogInAsAdmin();
            ScheduledJobsPage scheduledJobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = scheduledJobsPage.Nav_To_Cegid();
            try
            {
                CegidScheduledJobModal cegidScheduledJobModal = scheduledJobsCegidPage.NewCegidScheduledModal();
                cegidScheduledJobModal.FillFieldsModalCegidScheduledJob(name, fiscalEntites);
                scheduledJobsCegidPage.ResetFilter();
                scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, name);
                var addedName = scheduledJobsCegidPage.GetFirstCegidName();
                var isNameCorrect = addedName.Equals(name);
                Assert.IsTrue(isNameCorrect, "L'ajout ne fonctionne pas corretement.");
            }
            finally
            {
                scheduledJobsCegidPage.ResetFilter();
                scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, name);
                scheduledJobsCegidPage.DeleteFirstRow();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_Run()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobsPage.Nav_To_Cegid();
            scheduledJobsCegidPage.CLickOnStart();

            var testVisiblity = scheduledJobsCegidPage.CheckStatutIsVisible();

            //Assert
            Assert.IsTrue(testVisiblity, "les 3 status ne sont pas visibles");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_filter_ShowAll()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobsPage.Nav_To_Cegid();
            scheduledJobsCegidPage.ResetFilter();
            scheduledJobsCegidPage.PageSize("100");

            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowAll, true);
            var resultAll = scheduledJobsCegidPage.GetTotalResultsAllPages();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowActiveOnly, true);
            var resultsActivated = scheduledJobsCegidPage.GetTotalResultsAllPages();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowInactiveOnly, true);
            var resultsInactivated = scheduledJobsCegidPage.GetTotalResultsAllPages();

            //Assert
            Assert.AreEqual(resultsActivated + resultsInactivated, resultAll, "Les résultats ne sont pas mis à jour en fonction des filtres Show All");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_filter_CegidJobType()
        {
            string cegidJobTypes = "WRSAGEINVENTORY";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobsPage.Nav_To_Cegid();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.GegidJobTypes, cegidJobTypes);

            //Assert
            bool gegidJobTypesIsFiltred = scheduledJobsCegidPage.GegidJobTypesFiltred(cegidJobTypes);
            Assert.IsTrue(gegidJobTypesIsFiltred, "l'application de filtre par gegidJobTypes ne fonctionne pas correctement");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_AddNew()
        {
            string fileFlowName = "testName" + DateTime.Today.ToShortDateString();
            bool isActive = true;

            // Arrange
            var homePage = LogInAsAdmin();
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobscheduled.Nav_To_File_Flows();
            try
            {
                FileFlowsCreateModalPage fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
                fileFlowsCreateModalPage.FillFiled_CreateFileFlows(fileFlowName, CUSTOMER, isActive, null, CONVERTER);
                scheduledJobsTabFileFlowsPage.ResetFilter();
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlowName);
                var verifyName = scheduledJobsTabFileFlowsPage.GetNameFirstRow().Equals(fileFlowName);
                var verifyCustomer = scheduledJobsTabFileFlowsPage.GetFirstLineCustomer().Equals(CUSTOMER);
                var verifyConverterType = scheduledJobsTabFileFlowsPage.GetFirstLineConverterType().Equals(CONVERTER);
                var dataCorrect = verifyName && verifyCustomer && verifyConverterType;
                Assert.IsTrue(dataCorrect, "L'ajout d'un nouveau Bank ne fonctionne pas correctement!");
            }
            finally
            {
                scheduledJobsTabFileFlowsPage.ResetFilter();
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlowName);
                scheduledJobsTabFileFlowsPage.DeleteFirstFileFlow();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_AddNewValidator()
        {
            var homePage = LogInAsAdmin();
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobscheduled.Nav_To_File_Flows();
            FileFlowsCreateModalPage fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
            fileFlowsCreateModalPage.FillFiled_CreateFileFlows("", "", false, null, null, "");
            var verifyValidator = fileFlowsCreateModalPage.VerifyCreateValidators();
            Assert.IsTrue(verifyValidator, "Les messages Erreur ne sont pas tous affiché !");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_AddNewValidator()
        {
            // Arrange
            var homePage = LogInAsAdmin();
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobsPage.Nav_To_Cegid();
            CegidScheduledJobModal cegidScheduledJobModal = scheduledJobsCegidPage.NewCegidScheduledModal();
            //simulate empty cron expression
            cegidScheduledJobModal.FillFieldsModalCegidScheduledJob("", null, false, "");
            var verifyValidator = cegidScheduledJobModal.VerifyCreateValidators();
            Assert.IsTrue(verifyValidator, "Les validateurs n'apparaissent pas tous !");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_Delete()
        {
            // Prepare
            string cegidName = "CegidToDelete";
            string fiscalEntities = "NEWREST ESPAÑA S.L.";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage scheduledJobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage cegidPage = scheduledJobsPage.Nav_To_Cegid();
            cegidPage.ResetFilter();

            cegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, cegidName);

            try
            {
                CegidScheduledJobModal cegidCreateModal = cegidPage.NewCegidScheduledModal();
                cegidCreateModal.FillFieldsModalCegidScheduledJob(cegidName, fiscalEntities, true);

                cegidPage.ResetFilter();
                cegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, cegidName);
                Assert.AreEqual(cegidName, cegidPage.GetFirstCegidName(), "L'élément Cegid n'a pas été ajouté.");
            }
            finally
            {
                scheduledJobsPage.Nav_To_Cegid();
                cegidPage.ResetFilter();
                cegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, cegidName);
                cegidPage.DeleteFirstRow();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_Pagination()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsPage = homePage.GoToJobs_ScheduledJobs();
            var scheduledJobsCegidPage = jobsPage.Nav_To_Cegid();

            //PageSize 8
            scheduledJobsCegidPage.PageSize("8");
            int count = scheduledJobsCegidPage.GetAllNamesPageResult().Count ;
            Assert.IsTrue(count <= 8, "Pagination ne fonctionne pas..");

            //PageSize 16
            scheduledJobsCegidPage.PageSize("16");
            count = scheduledJobsCegidPage.GetAllNamesPageResult().Count;
            Assert.IsTrue(count <= 16, "Pagination ne fonctionne pas..");

            //PageSize 30
            scheduledJobsCegidPage.PageSize("30");
            count = scheduledJobsCegidPage.GetAllNamesPageResult().Count;
            Assert.IsTrue(count <= 30, "Pagination ne fonctionne pas..");

            //PageSize 50
            scheduledJobsCegidPage.PageSize("50");
            count = scheduledJobsCegidPage.GetAllNamesPageResult().Count;
            Assert.IsTrue(count <= 50, "Pagination ne fonctionne pas..");

            //PageSize 100
            scheduledJobsCegidPage.PageSize("100");
            count = scheduledJobsCegidPage.GetAllNamesPageResult().Count;
            Assert.IsTrue(count <= 100, "Pagination ne fonctionne pas..");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_DeleteValidator()
        {
            string cegidName = "CegidToDelete";
            string fiscalEntities = "NEWREST ESPAÑA S.L.";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage scheduledJobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage cegidPage = scheduledJobsPage.Nav_To_Cegid();
            cegidPage.ResetFilter();

            if (cegidPage.GetCountResultPaged() == 0)
            {
                // Ajouter un job Cegid pour tester la suppression
                CegidScheduledJobModal cegidCreateModal = cegidPage.NewCegidScheduledModal();
                cegidCreateModal.FillFieldsModalCegidScheduledJob(cegidName, fiscalEntities, true);

                cegidPage.ResetFilter();
                cegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, cegidName);
                Assert.AreEqual(cegidName, cegidPage.GetFirstCegidName(), "L'élément Cegid n'a pas été ajouté.");
            }
            string cegidNameToDelete = cegidPage.GetFirstCegidName();
            cegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, cegidNameToDelete);
            cegidPage.ClickDelete();
            bool popupOpened = cegidPage.CheckOpenedPopup();

            //Assert
            Assert.IsTrue(popupOpened, "La pop-up de confirmation n'est pas affichée.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_editPopUp()
        {
            string name = "CegidForAdd";
            string fiscalEntites = "NEWREST ESPAÑA S.L.";

            var homePage = LogInAsAdmin();
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobscheduled.Nav_To_Cegid();
            CegidScheduledJobModal cegidScheduledJobModal = scheduledJobsCegidPage.NewCegidScheduledModal();
            cegidScheduledJobModal.FillFieldsModalCegidScheduledJob(name, fiscalEntites);
            scheduledJobsCegidPage.ClickOnFirstRow();
            var isOpened = scheduledJobsCegidPage.CheckOpenedPopup();
            Assert.IsTrue(isOpened, "La popup ne s'affiche pas!");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Cegid_CronExpressionEditor()

        {
            var homePage = LogInAsAdmin();
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobsPage.Nav_To_Cegid();
            scheduledJobsCegidPage.ResetFilter();
            CegidScheduledJobModal cegidScheduledJobModal = scheduledJobsCegidPage.ClickFirstLine();
            cegidScheduledJobModal.EditorCronExpression();
            bool tabsCorrect = cegidScheduledJobModal.CheckTabsExistance();
            Assert.IsTrue(tabsCorrect, "Les onglets ne s'affichent pas!");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_edit()
        {
            string rnd = new Random().Next(1000, 9999).ToString();
            string name = "CegidForAdd";
            string fiscalEntites = "NEWREST ESPAÑA S.L.";
            var newName = "CegidForEdit" + rnd;
            
            var homePage = LogInAsAdmin();
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobscheduled.Nav_To_Cegid();
            try
            {
                CegidScheduledJobModal cegidAddScheduledJobModal = scheduledJobsCegidPage.NewCegidScheduledModal();
                cegidAddScheduledJobModal.FillFieldsModalCegidScheduledJob(name, fiscalEntites);
                scheduledJobsCegidPage.ResetFilter();
                scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, name);
                CegidScheduledJobModal cegidEditScheduledJobModal = scheduledJobsCegidPage.EditFirstRow();
                cegidEditScheduledJobModal.FillFieldsModalCegidScheduledJob(newName, fiscalEntites);
                scheduledJobsCegidPage.ResetFilter();
                scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, newName);
                var existedName = scheduledJobsCegidPage.GetFirstCegidName();
                Assert.AreEqual(existedName, newName, "Les infos n'ont pas été mises à jour");
            }
            finally
            {
                scheduledJobsCegidPage.DeleteFirstRow();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Bank_CronExpressionEditor()

        {
            var homePage = LogInAsAdmin();
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ClickFirstLine();
            scheduledJobsBankPage.EditorCronExpression();
            bool tabsCorrect = scheduledJobsBankPage.CheckTabsExistance();
            Assert.IsTrue(tabsCorrect, "Les onglets ne s'affichent pas!");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_filter_ResetFilter()
        {
            //Prepare
            string fileFlow = "Runfile_test_";
            string fileFlow2 = "Runfile_test_2";
            string FILE_FLOW_NAME = "FileFlow" + DateUtils.Now.ToString("dd/MM/yyyy") + "2";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobscheduled.Nav_To_File_Flows();
            try
            {
                scheduledJobsTabFileFlowsPage = jobscheduled.Nav_To_File_Flows();

                /* creat first file flow */
                var fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
                fileFlowsCreateModalPage.FillFiled_CreateFileFlows(fileFlow, CUSTOMER, true, null, CONVERTER);
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlow);
                Assert.IsTrue(scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage() >= 1, "Le File Flows n'est pas créé.");

                var defaultShow = scheduledJobsTabFileFlowsPage.GetShowFilterSelected();
                var defaultCustomersSelected = scheduledJobsTabFileFlowsPage.GetSelectedCustomersToFilter();
                var defaultConverterTypesSelected = scheduledJobsTabFileFlowsPage.GetConverterSelectedToFilter();
                var defaultItemsDisplayed = scheduledJobsTabFileFlowsPage.GetNumberItemsDisplayed();
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlow);
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.Converter, CONVERTER);
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.Customers, CUSTOMER);
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowAll, true);

                scheduledJobsTabFileFlowsPage.ResetFilter();

                var resultShowAfterReset = scheduledJobsTabFileFlowsPage.GetShowFilterSelected();
                var resultCustomersSelectedAfterReset = scheduledJobsTabFileFlowsPage.GetSelectedCustomersToFilter();
                var resultConverterTypesSelectedAfterReset = scheduledJobsTabFileFlowsPage.GetConverterSelectedToFilter();
                var resultItemsDisplayedAfterResult = scheduledJobsTabFileFlowsPage.GetNumberItemsDisplayed();

                //Assert
                Assert.AreEqual(defaultShow, resultShowAfterReset, "Le filter Show ne remet pas");
                Assert.AreEqual(defaultCustomersSelected.Count, resultCustomersSelectedAfterReset.Count, "Le filter CUSTOMERS ne remet pas");
                Assert.AreEqual(defaultConverterTypesSelected.Count, resultConverterTypesSelectedAfterReset.Count, "Le filter CONVERTER  ne remet pas");
                Assert.AreNotEqual(defaultItemsDisplayed, resultItemsDisplayedAfterResult, "Le filter  ne remet pas");
            }
            finally
            {
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlow);
                jobscheduled.DeleteFirstFileFlow();
            }
            
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_FileFlows_CronExpressionEdit()
        {
            string fileFlow = "Runfile_test";
            
            var homePage = LogInAsAdmin();
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobscheduled.Nav_To_File_Flows();
            try
            {
                var fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
                fileFlowsCreateModalPage.FillFiled_CreateFileFlows(fileFlow, CUSTOMER, true, null, CONVERTER);
                scheduledJobsTabFileFlowsPage.ResetFilter();
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlow);
                var isAdded = scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage() >= 1;
                Assert.IsTrue(isAdded, "Le File Flows n'est pas créé.");
                var oldFrequency = scheduledJobsTabFileFlowsPage.GetFirstRowFrequency();
                fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.EditFileFlow();
                fileFlowsCreateModalPage.ClickOnEditorButton();
                fileFlowsCreateModalPage.SelectTab(FileFlowsCreateModalPage.Tabs.MINUTES);
                fileFlowsCreateModalPage.EditMinutesCronExpression();
                fileFlowsCreateModalPage.SelectTab(FileFlowsCreateModalPage.Tabs.HOURS);
                fileFlowsCreateModalPage.EditHourCronExpression();
                scheduledJobsTabFileFlowsPage.ResetFilter();
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlow);
                var newFrequency = scheduledJobsTabFileFlowsPage.GetFirstRowFrequency();
                Assert.AreNotEqual(newFrequency, oldFrequency, "La modification du Cron expression ne fonctionne pas!");
            }
            finally
            {
                scheduledJobsTabFileFlowsPage.DeleteFirstFileFlow();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_filter_Converters()
        {
            //Prepare
            string fileFlow = "Runfile_test";
            string converterTypeToFilter = "AC_HR_txt";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            try
            {
                ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobsPage.Nav_To_File_Flows();
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.ShowAll, true);

                //create file
                var fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
                fileFlowsCreateModalPage.FillFiled_CreateFileFlows(fileFlow, CUSTOMER, true, null, CONVERTER);
                scheduledJobsTabFileFlowsPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.SearchByName, fileFlow);
                int totalResultRowsPage = scheduledJobsTabFileFlowsPage.GetTotalResultRowsPage();

                //Assert
                Assert.IsTrue(totalResultRowsPage >= 1, "Le File Flows n'est pas créé.");
                jobsPage.Filter(ScheduledJobsPage.FilterType.ConverterTypes, converterTypeToFilter);
                var verifyConverterFilter = jobsPage.CheckAllRowsIsSpecificConverters(converterTypeToFilter);
                Assert.IsTrue(verifyConverterFilter, "Le Filter par converters n'est pas fonctionnel ! ");

            }
            finally
            {
                //delete file
                jobsPage.DeleteFirstFileFlow();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_filter_ShowActiveOnly()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage scheduledJobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = scheduledJobsPage.Nav_To_Cegid();
            scheduledJobsCegidPage.ResetFilter();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowActiveOnly, true);
            scheduledJobsCegidPage.PageSize("100");
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowAll, true);
            var resultAll = scheduledJobsCegidPage.GetTotalResultsAllPages();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowActiveOnly, true);
            var resultsActivated = scheduledJobsCegidPage.GetTotalResultsAllPages();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowInactiveOnly, true);
            var resultsInactivated = scheduledJobsCegidPage.GetTotalResultsAllPages();
            Assert.AreEqual(resultAll - resultsInactivated, resultsActivated, "Les résultats ne sont pas mis à jour en fonction des filtres Show ActiveOnly");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_edit()
        {
            string nameBank = "testName" + DateTime.Now.ToString();
            string companyCode = "testCompanyCod";
            string sftpConnectionString = "testSFTPstring";
            string sftpFolder = "testSFTPFolderstring";
            string customer = TestContext.Properties["CustomerDelivery"].ToString();
            string converterType = "Default";
            string filenameToUpdate = "-FileUpdated" + new Random().Next().ToString();
            string converterTypeToUpdate = "JAF_csv";
            string companyCodeToUpdate = "testCompanyEdited";

            var homePage = LogInAsAdmin();
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobscheduled.Nav_To_Bank();
            jobscheduled.ResetFilter();
            if (scheduledJobsBankPage.GetTotalResultRowsPage() == 0 || (scheduledJobsBankPage.GetAllPageResultConverterType().All(item => item != converterType)))
                scheduledJobsBankPage.AddNewBank(nameBank, companyCode, sftpConnectionString, sftpFolder, true, customer, converterType);
            string defaultname = jobscheduled.GetNameFirstRow();
            var fileFlowEditModal = jobscheduled.SelectFirstRow();
            try
            {
                fileFlowEditModal.FillField_UpdateFileFlow(filenameToUpdate, converterTypeToUpdate, companyCodeToUpdate);
                scheduledJobsBankPage.ResetFilters();
                scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.SearchByName, filenameToUpdate);
                var numberOfRow = jobscheduled.GetTotalResultRowsPage() == 0;
                Assert.IsFalse(numberOfRow, "Le nom à modifier n'est pas enregistré");
                var verifyName = jobscheduled.GetNameFirstRow().Equals(filenameToUpdate);
                var verifyConverterType = jobscheduled.GetConverterTypeFirstRow().Equals(converterTypeToUpdate);
                var verifyCompanyCode = jobscheduled.GetCompanyCodeFirstRow().Equals(companyCodeToUpdate);
                var isEditCorrect = verifyName && verifyConverterType && verifyCompanyCode;
                Assert.IsTrue(isEditCorrect, "Les données ne sont pas modifiés correctement");
            }
            finally
            {
                fileFlowEditModal = jobscheduled.SelectFirstRow();
                fileFlowEditModal.FillField_UpdateFileFlow(defaultname, converterType, companyCode);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Bank_CronExpressionEdit()
        {
            var homePage = LogInAsAdmin();
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsBankPage scheduledJobsBankPage = jobsPage.Nav_To_Bank();
            scheduledJobsBankPage.ResetFilters();
            var name = scheduledJobsBankPage.GetFirstLineName();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.SearchByName, name);
            var oldFrequency = scheduledJobsBankPage.GetFirstFrequency();
            BankCreateModalPage bankCreateModalPage = scheduledJobsBankPage.EditBank();
            bankCreateModalPage.ClickOnEditorButton();
            bankCreateModalPage.SelectTab(BankCreateModalPage.Tabs.HOURS);
            bankCreateModalPage.EditHourCronExpression();
            bankCreateModalPage.SelectTab(BankCreateModalPage.Tabs.MINUTES);
            bankCreateModalPage.EditMinutesCronExpression();
            scheduledJobsBankPage.Filter(ScheduledJobsBankPage.FilterType.SearchByName, name);
            var newFrequency = scheduledJobsBankPage.GetFirstFrequency();
            Assert.AreNotEqual(newFrequency, oldFrequency, "La modification du Frequency ne fonctionne pas!");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_filter_ShowInactiveOnly()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobsPage.Nav_To_Cegid();
            scheduledJobsCegidPage.ResetFilter();
            scheduledJobsCegidPage.PageSize("100");
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowAll, true);
            var resultAll = scheduledJobsCegidPage.GetTotalResultsAllPages();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowActiveOnly, true);
            var resultsActivated = scheduledJobsCegidPage.GetTotalResultsAllPages();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowInactiveOnly, true);
            var resultsInactivated = scheduledJobsCegidPage.GetTotalResultsAllPages();
            Assert.AreEqual(resultAll - resultsActivated, resultsInactivated, "Les résultats ne sont pas mis à jour en fonction des filtres ShowInactiveOnly");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Cegid_CronExpressionEdit()
        {
            var homePage = LogInAsAdmin();
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsCegidPage scheduledJobsCegidPage = jobsPage.Nav_To_Cegid();
            scheduledJobsCegidPage.ResetFilter();
            var name = scheduledJobsCegidPage.GetFirstLineName();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, name);
            var oldFrequency = scheduledJobsCegidPage.GetFirstFrequency();
            CegidScheduledJobModal cegidScheduledJobModal = scheduledJobsCegidPage.EditFirstRow();
            cegidScheduledJobModal.ClickOnEditorButton();
            cegidScheduledJobModal.SelectTab(CegidScheduledJobModal.Tabs.HOURS);
            cegidScheduledJobModal.EditHourCronExpression();
            cegidScheduledJobModal.SelectTab(CegidScheduledJobModal.Tabs.MINUTES);
            cegidScheduledJobModal.EditMinutesCronExpression();
            scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, name);
            var newFrequency = scheduledJobsCegidPage.GetFirstFrequency();
            Assert.AreNotEqual(newFrequency, oldFrequency, "La modification du Frequency ne fonctionne pas!");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_index_Pagination()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobscheduledPage = homePage.GoToJobs_ScheduledJobs();
            jobscheduledPage.PageSize("8");
            Assert.IsTrue(jobscheduledPage.GetNameProvidersList().Count <= 8, "Paggination ne fonctionne pas..");
            jobscheduledPage.PageSize("16");
            Assert.IsTrue(jobscheduledPage.GetNameProvidersList().Count <= 16, "Paggination ne fonctionne pas..");
            jobscheduledPage.PageSize("30");
            Assert.IsTrue(jobscheduledPage.GetNameProvidersList().Count <= 30, "Paggination ne fonctionne pas..");
            jobscheduledPage.PageSize("50");
            Assert.IsTrue(jobscheduledPage.GetNameProvidersList().Count <= 50, "Paggination ne fonctionne pas..");
            jobscheduledPage.PageSize("100");
            Assert.IsTrue(jobscheduledPage.GetNameProvidersList().Count <= 100, "Paggination ne fonctionne pas..");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Specific_CronExpressionEdit()
        {
            var homePage = LogInAsAdmin();
            var jobscheduled = homePage.GoToJobs_ScheduledJobs();
            ScheduledSpecificPage scheduledSpecificPage = jobscheduled.Nav_To_Specific();
            scheduledSpecificPage.ResetFilter();
            var SpecificName = scheduledSpecificPage.GetFirstLineName();
            scheduledSpecificPage.Filter(ScheduledSpecificPage.FilterType.SearchByName, SpecificName);
            var oldFrequency = scheduledSpecificPage.GetFirstFrequency();
            SpecificEditModal specificEditModal = scheduledSpecificPage.EditSpecific();
            specificEditModal.ClickOnEditorButton();
            specificEditModal.SelectTab(SpecificEditModal.Tabs.HOURS);
            specificEditModal.EditHourCronExpression();
            specificEditModal.SelectTab(SpecificEditModal.Tabs.MINUTES);
            specificEditModal.EditMinutesCronExpression();
            scheduledSpecificPage.Filter(ScheduledSpecificPage.FilterType.SearchByName, SpecificName);
            var newFrequency = scheduledSpecificPage.GetFirstFrequency();
            Assert.AreNotEqual(newFrequency, oldFrequency, "La modification du Frequency ne fonctionne pas!");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_Specific_CronExpressionEditor()

        {
            var homePage = LogInAsAdmin();
            ScheduledJobsPage jobsPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledSpecificPage scheduledSpecificPage = jobsPage.Nav_To_Specific();
            scheduledSpecificPage.ResetFilter();
            var specificName = scheduledSpecificPage.GetFirstLineName();
            scheduledSpecificPage.Filter(ScheduledSpecificPage.FilterType.SearchByName, specificName);
            SpecificEditModal specificEditModal = scheduledSpecificPage.EditSpecific();
            specificEditModal.ClickOnEditorButton();
            bool tabsCorrect = specificEditModal.CheckTabsExistance();
            Assert.IsTrue(tabsCorrect, "Les onglets ne s'affichent pas!");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SCHED_CEGID_RunValidator()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobScheduledPage = homePage.GoToJobs_ScheduledJobs();
            ScheduledJobsTabFileFlowsPage scheduledJobsTabFileFlowsPage = jobScheduledPage.Nav_To_File_Flows();
            var fileFlowsCreateModalPage = scheduledJobsTabFileFlowsPage.OpenModalCreateFileFLows();
            FileFlowsCreateModalPage fileFlowsPage = fileFlowsCreateModalPage.FillFiled_CreateFileFlows(FILE_FLOW_NAME, CUSTOMER, true);
            jobScheduledPage.RunJob();
            var isShowed = jobScheduledPage.CheckOpenedConfirmRunJobModal();

            //Assert
            Assert.IsTrue(isShowed, "La pop-up n'a pas apparu!");
        }

    }

}



