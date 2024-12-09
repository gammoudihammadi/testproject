using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Edi;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Catalogs;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.Jobs
{
    [TestClass]
    public class SettingsTests : TestBase
    {
        private const int _timeout = 600000;
        private readonly string FILEFLOWPROVIDER = "FlowManager";
        private readonly string FOLDERNAMER = "FlowFolder";
        private readonly string FILEFLOWTYPE = "FlowCategory";
        private readonly string CUSTOMER = "3D COMM";
        private readonly string CONVERTERTYPE = "AC_HR_txt";
        private readonly string EMAIL = "test.auto@newrest.eu";
        private readonly string FILEFLOW = "FileFlow";
        private readonly string SPECIFIC = "Specific";
        private readonly string CONVERTER = "Converter";
        private readonly string NOTIFTYPECUSTOMER = "Customer";
        private readonly string SAGE = "Sage";
        private readonly string CUSTOMEREMAIL = "Customer@newrest.eu";
        private readonly string CONVERTEREMAIL = "Converter@newrest.eu";
        private readonly string FILEFLOWEMAIL = "FileFlow@newrest.eu";
        private readonly string SAGEEMAIL = "Sage@newrest.eu";
        private readonly string SPECIFICEMAIL = "Specific@newrest.eu";
        private readonly string FILE_FLOW_NAME = "FileFlow" + DateUtils.Now.ToString("dd/MM/yyyy");
        private readonly string CEGIDSCHEDULED = "CegidNewRest2024";
        private readonly string SPECIFICSCHEDULE = "SpecificNewRest2024";
        private readonly string CUSTOMERCHAMPS = "//*[@id=\"SelectedCustomer-Group\"]/label";



        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_FilteFlowTypes_filter_SearchByName()
        {
            // Prepare
            string filteFlowType = "fileflowtype-" + new Random().Next().ToString();
            string foldername = "foldername-" + new Random().Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsetting = homePage.GoToJobs_Settings();
            var fileflowtype = jobsetting.GoToFILEFLOWPAGE();
            var modalcreatefileflowtype = fileflowtype.CreateModelFileFlowType();
            fileflowtype = modalcreatefileflowtype.FillField_CreateNewFileFlowType(filteFlowType, foldername);
            fileflowtype.ResetFilter();
            fileflowtype.Filter(FileFlowTypePage.FilterType.Search, filteFlowType);

            int nombreRowFileFlowType =fileflowtype.NombreRowFileFlowType();

            //Assert
            Assert.IsTrue(nombreRowFileFlowType >=1, "Les lignes ayant un Name correspondant à la recherche n'apparaitre pas.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_FileFlowTypes_AddNew()
        {
            // Prepare
            string FilteFlowType = "fileflowtype-" + new Random().Next().ToString();
            string foldername = "foldername-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            var fileflowtype = jobsetting.GoToFILEFLOWPAGE();
            try
            {
                var modalcreatefileflowtype = fileflowtype.CreateModelFileFlowType();
                fileflowtype = modalcreatefileflowtype.FillField_CreateNewFileFlowType(FilteFlowType, foldername);
                fileflowtype.ResetFilter();
                fileflowtype.Filter(FileFlowTypePage.FilterType.Search, FilteFlowType);
                bool filteFlowTypeName = fileflowtype.FindRowFileFlowTypeName(FilteFlowType);
                bool filteFlowTypeFolder = fileflowtype.FindRowFileFlowTypeFolder(foldername);
                Assert.IsTrue(filteFlowTypeName && filteFlowTypeFolder, "filteFlowTypeName et filteFlowTypeFolder n'apparaitre pas.");
            }
            finally
            {
                fileflowtype.DeleteFirstRow();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_FileFlowTypes_AddNewValidator()
        {
            
            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            var fileflowtype = jobsetting.GoToFILEFLOWPAGE();
            var modalcreatefileflowtype = fileflowtype.CreateModelFileFlowType();
            modalcreatefileflowtype.FillField_CreateNewFileFlowTypeWithoutData();
            var Validators =  modalcreatefileflowtype.IsValidatorsExists();
            Assert.IsTrue(Validators, "Les validators n'apparaissent pas pour les deux champs");


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_AddNew_ConverterJob()
        {
            Random rnd = new Random();
            var email = CUSTOMEREMAIL + rnd.Next().ToString();
            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            var jobnotification = jobsetting.GoTo_JOB_NOTIFICATION();
            try
            {
                var convertersjobnotificationmodel = jobnotification.AddNewConverterJobNotification();
                jobnotification = convertersjobnotificationmodel.FillField_AddNewCONVERTERSJOBNOTIFICATION(email, CUSTOMER);
                jobnotification.Filter(JobNotifications.FilterType.Search, email);
                jobnotification.Filter(JobNotifications.FilterType.NotificationTypes, CONVERTER);
                Assert.AreNotEqual(0,jobnotification.CountJobs(), "La notification n' apparait pas dans la liste avec les infos renseignées.");
            }
            finally
            {
                jobnotification.DeleteFirstJob();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_AddNew_ConverterJobValidator()
        {
            
            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            var jobnotification = jobsetting.GoTo_JOB_NOTIFICATION();
            var convertersjobnotificationmodel = jobnotification.AddNewConverterJobNotification();
            convertersjobnotificationmodel.FillField_AddNewCONVERTERSJOBNOTIFICATION("", "");
            Assert.IsTrue(convertersjobnotificationmodel.IsValidatorsExists(), "Les validators n'apparaissent pas");


        }

        [TestMethod]
        public void JB_SETT_FileFlowTypes_Delete()
        {
            // Prepare
            string FilteFlowType = "fileflowtype-" + new Random().Next().ToString();
            string foldername = "foldername-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var jobsetting = homePage.GoToJobs_Settings();
            var fileflowtype = jobsetting.GoToFILEFLOWPAGE();
            var modalcreatefileflowtype = fileflowtype.CreateModelFileFlowType();
            fileflowtype = modalcreatefileflowtype.FillField_CreateNewFileFlowType(FilteFlowType, foldername);
            fileflowtype.ResetFilter();
            fileflowtype.PageSize("100");
            var initialCount = fileflowtype.NombreRowFileFlowType();
            fileflowtype.Filter(FileFlowTypePage.FilterType.Search, FilteFlowType);
            fileflowtype.DeleteFirstRow();
            fileflowtype.ResetFilter();
            fileflowtype.PageSize("100");
            var finalCount = fileflowtype.NombreRowFileFlowType();

            Assert.AreNotEqual(finalCount, initialCount, "FilteFlowTypeName n'est pas supprimé.");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_FileFlowTypes_Edit()
        {
            // Prepare
            string FilteFlowType1 = "fileflowtype-" + new Random().Next().ToString();
            string foldername1 = "foldername-" + new Random().Next().ToString();

            string FilteFlowType2 = "fileflowtype-" + new Random().Next().ToString();
            string foldername2 = "foldername-" + new Random().Next().ToString();

            //Arrange
            
            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            var fileflowtype = jobsetting.GoToFILEFLOWPAGE();
            try
            {
                var modalcreatefileflowtype = fileflowtype.CreateModelFileFlowType();
                fileflowtype = modalcreatefileflowtype.FillField_CreateNewFileFlowType(FilteFlowType1, foldername1);
                fileflowtype.ResetFilter();
                fileflowtype.Filter(FileFlowTypePage.FilterType.Search, FilteFlowType1);
                FileFlowProvidersEditModal fileFlowProvidersEditModal = fileflowtype.ClickFirstRow();
                fileflowtype = fileFlowProvidersEditModal.FillFields_UpdateFileFlow(FilteFlowType2, foldername2);
                fileflowtype.Filter(FileFlowTypePage.FilterType.Search, FilteFlowType2);
                bool filteFlowTypeName = fileflowtype.FindRowFileFlowTypeName(FilteFlowType2);
                bool filteFlowTypeFolder = fileflowtype.FindRowFileFlowTypeFolder(foldername2);
                Assert.IsTrue(filteFlowTypeName && filteFlowTypeFolder, "filteFlowTypeName et filteFlowTypeFolder n'apparaitre pas.");
            }
            finally
            {
                fileflowtype.DeleteFirstRow();
                fileflowtype.ResetFilter();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]

        public void JB_SETT_index_Edit()
        {
            string filenameToUpdate = " - FileUpdated";
            string folderNameToUpdate = " - FolderUpdated";
            var OldFileName = "";
            var OldFolderName = "";
            //Arrange

                var homePage = LogInAsAdmin();
                var jobsettingsPage = homePage.GoToJobs_Settings();
            try
            {
                jobsettingsPage.ResetFilter();
                OldFileName = jobsettingsPage.GetNameFirstRow();
                OldFolderName = jobsettingsPage.GetFolderFirstRow();
                var fileFlowProvidersEditModal = jobsettingsPage.SelectFirstRow();
                fileFlowProvidersEditModal.FillField_UpdateFileFlowProviders(filenameToUpdate, folderNameToUpdate);
                var NewUpdatedName = OldFileName + filenameToUpdate;
                var NewUpdateFolder = OldFolderName + folderNameToUpdate;
                jobsettingsPage.Filter(SettingsPage.FilterType.Search, NewUpdatedName);
                var NewFileName = jobsettingsPage.GetNameFirstRow();
                var NewFolderName = jobsettingsPage.GetFolderFirstRow();

                Assert.AreEqual(NewUpdatedName , NewFileName, "filteFlowTypeName n'apparaitre pas.");
                Assert.AreEqual(NewUpdateFolder, NewFolderName, "filteFlowTypeFolder n'apparaitre pas.");
            }
            finally 
            {
                var fileFlowProvidersEditModal = jobsettingsPage.SelectFirstRow();
                fileFlowProvidersEditModal.FillFields_UpdateFileFlow(OldFileName, OldFolderName);

            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_index_AddNew()
        {
            // Prepare
            string filteFlowProvider = "fileflowprovider-" + new Random().Next().ToString();
            string foldername = "foldername-" + new Random().Next().ToString();
            //Arrange
            
            var homePage = LogInAsAdmin();
            var jobsettingsPage = homePage.GoToJobs_Settings();
            var fileflowprovider = jobsettingsPage.GoToFILEFLOWPROVIDERPAGE();
            try
            {
                var modalcreatefileflowtype = fileflowprovider.CreateModelFileFlowProvider();
                fileflowprovider = modalcreatefileflowtype.FillField_CreateNewFileFlowProvider(filteFlowProvider, foldername);
                fileflowprovider.ResetFilter();
                fileflowprovider.Filter(FileFlowProviderPage.FilterType.Search, filteFlowProvider);
                bool filteFlowProviderName = fileflowprovider.FindRowFileFlowProviderName(filteFlowProvider);
                bool filteFlowProviderFolder = fileflowprovider.FindRowFileFlowProviderFolder(foldername);
                Assert.IsTrue(filteFlowProviderName && filteFlowProviderFolder, "filteFlowProviderName et filteFlowProviderFolder n'apparaitre pas.");
            }
            finally
            {
                jobsettingsPage.DeleteElementFileFlowProviders(filteFlowProvider);
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_index_filter_SearchByName()
        {
            string filteFlowProvider = "fileflowprovider-" + new Random().Next().ToString();
            string foldername = "foldername-" + new Random().Next().ToString();

            var homePage = LogInAsAdmin();
            var jobsettingsPage = homePage.GoToJobs_Settings();
            var fileflowprovider = jobsettingsPage.GoToFILEFLOWPROVIDERPAGE();
            try 
            {
            fileflowprovider.ResetFilter();
            var modalcreatefileflowprovider = fileflowprovider.CreateModelFileFlowProvider();
            fileflowprovider = modalcreatefileflowprovider.FillField_CreateNewFileFlowProvider(filteFlowProvider, foldername);
            fileflowprovider.ResetFilter();
            fileflowprovider.Filter(FileFlowProviderPage.FilterType.Search, filteFlowProvider);
            Assert.IsTrue(fileflowprovider.FindRowFileFlowProviderName(filteFlowProvider), "Les lignes ayant un Name correspondant à la recherche n'apparaitre pas.");
            }
            finally
            {
              fileflowprovider.DeleteFirstRow();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_index_Delete()
        {
            // Prepare
            string filteFlowProvider = "fileflowprovider-" + new Random().Next().ToString();
            string foldername = "foldername-" + new Random().Next().ToString();
           
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsettingsPage = homePage.GoToJobs_Settings();
            jobsettingsPage.ResetFilter();

            //create
            var modalcreatefileflowtype = jobsettingsPage.CreateModelFileFlowProviders();
            modalcreatefileflowtype.FillField_CreateNewFileFlowProviders(filteFlowProvider, foldername);

            //Delete
            jobsettingsPage.Filter(SettingsPage.FilterType.Search, filteFlowProvider);
            jobsettingsPage.DeleteElementFileFlowProviders(filteFlowProvider);
            jobsettingsPage.ResetFilter();
            var listElemnetsBeforeAfter = jobsettingsPage.GetElementsNamesFileFlowProviders();

            Assert.IsFalse(listElemnetsBeforeAfter.Contains(filteFlowProvider), "Delete Was Failed");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_index_AddNewValidators()
        {
            // Prepare
            string FilteFlowProvider = " ";
            string foldername = " ";
            
            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            FileFlowProviderPage fileFlowProviderPage = jobsetting.GoToFILEFLOWPROVIDERPAGE();
            FileFlowProviderCreateModalPage fileFlowProvidersCreateModel = fileFlowProviderPage.CreateModelFileFlowProvider();
            fileFlowProvidersCreateModel.FillField_CreateNewFileFlowProvider(FilteFlowProvider, foldername);
            Assert.IsTrue(fileFlowProvidersCreateModel.IsNameRequiredMessageExist(), "Le validateur du champ Name n'apparaît pas.");
            Assert.IsTrue(fileFlowProvidersCreateModel.IsFolderNameRequiredMessageExist(), "Le validateur du champ Folder Name n'apparaît pas.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_Filter_email()
        {
            //Prepare
            var email = "imen@newrest.com";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsetting = homePage.GoToJobs_Settings();
            var jobnotification = jobsetting.GoTo_JOB_NOTIFICATION();
            var convertersjobnotificationmodel = jobnotification.AddNewConverterJobNotification();
            convertersjobnotificationmodel.FillField_AddNewCONVERTERSJOBNOTIFICATION(email);
            jobnotification.FilterByEmail(email);
            var emaillist = jobnotification.GetEmailList();
            bool emailFound = emaillist.Any(item => item.Contains(email));

            //Assert
            Assert.IsTrue(emailFound, $"The email '{email}'n'a pas été trouvé dans la liste.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_Delete()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsetting = homePage.GoToJobs_Settings();
            var jobnotification = jobsetting.GoTo_JOB_NOTIFICATION();
            jobnotification.PageSize("100");
            var emailListBeforeDelete = jobnotification.GetEmailList();

            jobnotification.DeleteFirstJob();
            var emailListAfterDelete = jobnotification.GetEmailList();

            //Assert
            Assert.AreNotEqual(emailListBeforeDelete, emailListAfterDelete, $"Nombre de notification avant suppression est '{emailListBeforeDelete}' , et le Nombre de notif apres suppression est '{emailListAfterDelete}'  ===>  La notification n'est pas supprimé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_filters_Customer()
        {
            //Prepare
            string customer = TestContext.Properties["CustomerJob"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var jobsetting = homePage.GoToJobs_Settings();
            var jobnotification = jobsetting.GoTo_JOB_NOTIFICATION();
            jobnotification.PageSize("8");
            jobnotification.FilterByCutomer(customer);
            bool allJobsBelongToCustomer = jobnotification.VerifyAllJobsBelongToCustomer(customer);

            //Assert
            Assert.IsTrue(allJobsBelongToCustomer, $"Le filtre customer pour '{customer}' ne fonctionne pas.");
        }

        [TestMethod]
        public void JobNotificationFilters_SearchPattern()
        {
            var email = "marouane@gmail.com";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var jobsetting = homePage.GoToJobs_Settings();
            var jobnotification = jobsetting.GoTo_JOB_NOTIFICATION();
            var jobsCount = jobnotification.CountJobs();
            jobnotification.FilterByEmail(email);
            var jobsSecoundCount = jobnotification.CountJobs();

            Assert.AreNotEqual(jobsCount, jobsSecoundCount, "Le filtre customer ne fonctionne pas.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_Converters_filter_Customers()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string converterType = "AC_HR_txt";

            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            var settingsTabConvertersPage = jobsetting.GoTo_CONVERTERS();
            try
            {
                var convertersCreateModalPage = settingsTabConvertersPage.AddNewConverter();
                var converterDetailsPage = convertersCreateModalPage.FillFiled_CreateConverters(customer, converterType);
                settingsTabConvertersPage = converterDetailsPage.BackToList();
                settingsTabConvertersPage.ResetFilter();
                settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.Customers, customer);
                var isFilterCorrect = settingsTabConvertersPage.GetCustomersFiltred().All(cc => cc.Equals(customer));
                Assert.IsTrue(isFilterCorrect, "l'application de filtre par customer ne fonctionne pas correctement");
            }
            finally
            {
                settingsTabConvertersPage.deleteFirstConverters();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_FileFlowTypes_pagination()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsetting = homePage.GoToJobs_Settings();

            //Appliquer PageSize 8
            jobsetting.PageSize("8");
            var count = jobsetting.GetFileFlowProvidersList().Count;
            //Assert
            Assert.IsTrue(count <= 8, "Pagination ne fonctionne pas..");

            //Appliquer PageSize 16
            jobsetting.PageSize("16");
            count = jobsetting.GetFileFlowProvidersList().Count;
            Assert.IsTrue(count <= 16, "Pagination ne fonctionne pas..");

            //Appliquer PageSize 30
            jobsetting.PageSize("30");
            count = jobsetting.GetFileFlowProvidersList().Count;
            Assert.IsTrue(count <= 30, "Pagination ne fonctionne pas..");

            //Appliquer PageSize 50
            jobsetting.PageSize("50");
            count = jobsetting.GetFileFlowProvidersList().Count;
            Assert.IsTrue(count <= 50, "Pagination ne fonctionne pas..");

            //Appliquer PageSize 100
            jobsetting.PageSize("100");
            count = jobsetting.GetFileFlowProvidersList().Count;
            Assert.IsTrue(count <= 100, "Pagination ne fonctionne pas..");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_Converters_filter_ConverterTypes()

        {
            //prepare 
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string converterType = "AC_HR_txt";

            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            var settingsTabConvertersPage = jobsetting.GoTo_CONVERTERS();
            try
            {
                var convertersCreateModalPage = settingsTabConvertersPage.AddNewConverter();
                var converterDetailsPage = convertersCreateModalPage.FillFiled_CreateConverters(customer, converterType);
                settingsTabConvertersPage = converterDetailsPage.BackToList();
                settingsTabConvertersPage.ResetFilter();
                settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.ConverterType, converterType);
                var isFilterCorrect = settingsTabConvertersPage.GetConverterTypeFiltred().All(con => con.Equals(converterType));
                Assert.IsTrue(isFilterCorrect, "l'application de filtre par converter Types ne fonctionne pas correctement");
            }
            finally
            {
                settingsTabConvertersPage.deleteFirstConverters();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_index_reset_filter()
        {
            // Prepare
            string filteFlowProvider = "fileflowprovider-" + new Random().Next().ToString();
            string foldername = "foldername-" + new Random().Next().ToString();
            string expectedTitle = "Fileflow providers";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsettingsPage = homePage.GoToJobs_Settings();
            var fileflowprovider = jobsettingsPage.GoToFILEFLOWPROVIDERPAGE();
            var modalcreatefileflowtype = fileflowprovider.CreateModelFileFlowProvider();
            fileflowprovider = modalcreatefileflowtype.FillField_CreateNewFileFlowProvider(filteFlowProvider, foldername);
            fileflowprovider.ResetFilter();
            fileflowprovider.Filter(FileFlowProviderPage.FilterType.Search, filteFlowProvider);
            fileflowprovider.ResetFilter();
            string actualTitle = fileflowprovider.GetFileflowProvidersTitle();
            
            //Assert
            Assert.AreEqual(expectedTitle, actualTitle, "Le titre de la page ne correspond pas au titre attendu 'Fileflow Providers'.");
            string inputValue = fileflowprovider.GetInputSearchValue();
            Assert.IsTrue(string.IsNullOrEmpty(inputValue), "Le champ de recherche n'est pas vide.");
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_FileFlowTypes_ResetFilter()
        {
            //Prepare
            string FilteFlowType = "fileflowtype-" + new Random().Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsetting = homePage.GoToJobs_Settings();
            var fileflowtype = jobsetting.GoToFILEFLOWPAGE();
            fileflowtype.PageSize("100");
            int countBeforeReset = fileflowtype.GetFileFlowTypesList().Count;

            fileflowtype.Filter(FileFlowTypePage.FilterType.Search, FilteFlowType);
            fileflowtype.ResetFilter();
            fileflowtype.PageSize("100");
            int countAfterReset = fileflowtype.GetFileFlowTypesList().Count;

            //Assert
            Assert.AreEqual(countBeforeReset, countAfterReset, "Reset Filter has failed");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_Converters_ResetFilter()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string converterType = "AC_HR_txt";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsetting = homePage.GoToJobs_Settings();
            var settingsTabConvertersPage = jobsetting.GoTo_CONVERTERS();
            try
            {
                //create a converter
                var convertersCreateModalPage = settingsTabConvertersPage.AddNewConverter();
                var converterDetailsPage = convertersCreateModalPage.FillFiled_CreateConverters(customer, converterType);
                settingsTabConvertersPage = converterDetailsPage.BackToList();

                string oldvalue = settingsTabConvertersPage.GetInputSearchCustomersValue();
                settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.Customers, customer);
                settingsTabConvertersPage.ResetFilter();
                string afterresetvalue = settingsTabConvertersPage.GetInputSearchCustomersValue();
                string actualTitle = settingsTabConvertersPage.GetConvertersTitle();
                string expectedTitle = "Converters";

                //Assert
                Assert.AreEqual(expectedTitle, actualTitle, "Le titre de la page ne correspond pas au titre attendu 'Converters'.");
                Assert.AreEqual(oldvalue, afterresetvalue, "Les filtres ajoutés ne sont pas réinitialisés après reset.");
            }
            finally
            {
                //delete
                settingsTabConvertersPage.deleteFirstConverters();
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_Converters_pagination()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            SettingsPage jobsetting = homePage.GoToJobs_Settings();
            SettingsTabConvertersPage settingsTabConvertersPage = jobsetting.GoTo_CONVERTERS();
            settingsTabConvertersPage.PageSize("8");
            Assert.IsTrue(settingsTabConvertersPage.GetCustomersFiltred().Count <= 8, "Paggination ne fonctionne pas..");
            jobsetting.PageSize("16");
            Assert.IsTrue(settingsTabConvertersPage.GetCustomersFiltred().Count <= 16, "Paggination ne fonctionne pas..");
            jobsetting.PageSize("30");
            Assert.IsTrue(settingsTabConvertersPage.GetCustomersFiltred().Count <= 30, "Paggination ne fonctionne pas..");
            jobsetting.PageSize("50");
            Assert.IsTrue(settingsTabConvertersPage.GetCustomersFiltred().Count <= 50, "Paggination ne fonctionne pas..");
            jobsetting.PageSize("100");
            Assert.IsTrue(settingsTabConvertersPage.GetCustomersFiltred().Count <= 100, "Paggination ne fonctionne pas..");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_Converters_Edit()
        {
            //Prepare
            string converterType = "AACA_XLSX";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            SettingsPage jobsetting = homePage.GoToJobs_Settings();
            var settingsTabConvertersPage = jobsetting.GoTo_CONVERTERS();
            ConverterDetailsPage converterDetailsPage = settingsTabConvertersPage.SelectFirstRow();
            converterDetailsPage.FillFields_EditConverters(CUSTOMER, converterType);
            converterDetailsPage.BackToList();
            settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.Customers, CUSTOMER);
            Assert.IsTrue(settingsTabConvertersPage.GetCustomersFiltred().All(cc => cc == CUSTOMER), "l'application de filtre par customer ne fonctionne pas correctement");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_Converters_Delete()
        {
            //prepare 
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string converterType = "AC_HR_txt";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsetting = homePage.GoToJobs_Settings();
            SettingsTabConvertersPage settingsTabConvertersPage = jobsetting.GoTo_CONVERTERS();
            ConvertersCreateModalPage convertersCreateModalPage = settingsTabConvertersPage.AddNewConverter();
            ConverterDetailsPage converterDetailsPage = convertersCreateModalPage.FillFiled_CreateConverters(customer, converterType);
            settingsTabConvertersPage = converterDetailsPage.BackToList();
            settingsTabConvertersPage.ResetFilter();
            settingsTabConvertersPage.PageSize("100");
            var customersBeforeDelete = settingsTabConvertersPage.GetCustomersFiltred().Count;
            settingsTabConvertersPage.deleteFirstConverters();
            //settingsTabConvertersPage.DeleteFirstRow();
            settingsTabConvertersPage.ResetFilter();
            settingsTabConvertersPage.PageSize("100");
            var customersAfterDelete = settingsTabConvertersPage.GetCustomersFiltred().Count;

            //Assert
            Assert.IsTrue(customersBeforeDelete > customersAfterDelete, "converter n'est pas supprimé.");
        }



        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void JB_SETT_index_CreateFileFlowProviders()
        {
            //Arrange
            
            var homePage = LogInAsAdmin();
            var jobsettingsPage = homePage.GoToJobs_Settings();
            var fileflowprovider = jobsettingsPage.GoToFILEFLOWPROVIDERPAGE();
            fileflowprovider.Filter(FileFlowProviderPage.FilterType.Search, FILEFLOWPROVIDER);

            if (fileflowprovider.NombreRowFileFlowProvider() == 0)
            {
                var modalcreatefileflowtype = fileflowprovider.CreateModelFileFlowProvider();
                fileflowprovider = modalcreatefileflowtype.FillField_CreateNewFileFlowProvider(FILEFLOWPROVIDER, FOLDERNAMER);
                fileflowprovider.ResetFilter();
                fileflowprovider.Filter(FileFlowProviderPage.FilterType.Search, FILEFLOWPROVIDER);
            }

                bool filteFlowProviderName = fileflowprovider.FindRowFileFlowProviderName(FILEFLOWPROVIDER);
                bool filteFlowProviderFolder = fileflowprovider.FindRowFileFlowProviderFolder(FOLDERNAMER);
                Assert.IsTrue(filteFlowProviderName && filteFlowProviderFolder, "filteFlowProviderName et filteFlowProviderFolder n'apparaitre pas.");
          
        }

        [TestMethod]
        [Priority(1)]
        [Timeout(_timeout)]
        public void JB_SETT_FileFlowTypes_CreateFileFlowTypeForTests()
        {
            //Arrange
            
            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            var fileflowtype = jobsetting.GoToFILEFLOWPAGE();
            fileflowtype.Filter(FileFlowTypePage.FilterType.Search, FILEFLOWTYPE);

            if (fileflowtype.NombreRowFileFlowType() == 0)
            {
                var modalcreatefileflowtype = fileflowtype.CreateModelFileFlowType();
                fileflowtype = modalcreatefileflowtype.FillField_CreateNewFileFlowType(FILEFLOWTYPE, FOLDERNAMER);
                fileflowtype.ResetFilter();
                fileflowtype.Filter(FileFlowTypePage.FilterType.Search, FILEFLOWTYPE);
                bool filteFlowTypeName = fileflowtype.FindRowFileFlowTypeName(FILEFLOWTYPE);
                bool filteFlowTypeFolder = fileflowtype.FindRowFileFlowTypeFolder(FOLDERNAMER);
                Assert.IsTrue(filteFlowTypeName && filteFlowTypeFolder, "filteFlowTypeName et filteFlowTypeFolder n'apparaitre pas.");
            }
            else
            {
                bool filteFlowTypeName = fileflowtype.FindRowFileFlowTypeName(FILEFLOWTYPE);
                bool filteFlowTypeFolder = fileflowtype.FindRowFileFlowTypeFolder(FOLDERNAMER);
                Assert.IsTrue(filteFlowTypeName && filteFlowTypeFolder, "filteFlowTypeName et filteFlowTypeFolder n'apparaitre pas.");
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_index_pagination()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobsetting = homePage.GoToJobs_Settings();
            jobsetting.ResetFilter();
            var fileflowprovider1 = jobsetting.GetFileFlowProvidersList()[0];
            var numberOfFileFlowProvider = jobsetting.GetFileFlowProvidersList().Count;
            Assert.IsTrue(numberOfFileFlowProvider == 8, "error");
            jobsetting.GoToPageTwo();
            var fileflowprovider2 = jobsetting.GetFileFlowProvidersList()[0];
            Assert.AreNotEqual(fileflowprovider1, fileflowprovider2, "page changing failed");
            jobsetting.PageSize("100");
            numberOfFileFlowProvider = jobsetting.GetFileFlowProvidersList().Count;
            Assert.IsTrue(numberOfFileFlowProvider > 8, "error");
        }

        [TestMethod]
        [Priority(2)]
        [Timeout(_timeout)]
        public void JB_SETT_Converters_CreateConverterForTests()
        {
            //Arrange
            
            var homePage = LogInAsAdmin();
            SettingsPage jobsetting = homePage.GoToJobs_Settings();
            SettingsTabConvertersPage settingsTabConvertersPage = jobsetting.GoTo_CONVERTERS();
            settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.Customers, CUSTOMER);
            settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.ConverterType, CONVERTERTYPE);

            // purge à la râche
            /* settingsTabConvertersPage.PageSize("100");
            for (int i = 0; i <  100- 1; i++)
            {
                settingsTabConvertersPage.DeleteFirstRow();
            } */


            if (settingsTabConvertersPage.GetCustomersFiltred().Count == 0)
            {
                ConvertersCreateModalPage convertersCreateModalPage = settingsTabConvertersPage.AddNewConverter();
                var converterDetailsPage = convertersCreateModalPage.FillFiled_CreateConverters(CUSTOMER, CONVERTERTYPE);
                settingsTabConvertersPage = converterDetailsPage.BackToList();
                settingsTabConvertersPage.ResetFilter();
                settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.Customers, CUSTOMER);
                settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.ConverterType, CONVERTERTYPE);
            }
            Assert.AreNotEqual(0, settingsTabConvertersPage.GetCustomersFiltred().Count, "customer et converter type n'apparaitre pas.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_AddNew_FileFlowJob()
        {
            Random rnd = new Random();
            var email = CUSTOMEREMAIL + rnd.Next().ToString();

            var homePage = LogInAsAdmin();
            var jobNotifications = homePage.GoToJobs_Settings().GoTo_JOB_NOTIFICATION();
            try {
                var fileFlowModal = jobNotifications.AddNewFileFlowJobNotification();
                fileFlowModal.FillFieldAddNewFileFlowJobNotification(email, CUSTOMER);
                jobNotifications.Filter(JobNotifications.FilterType.Search, email);
                var exist = jobNotifications.FindRowEmail(email);
                Assert.IsTrue(exist, "File flow job notification n'a pas été ajouté.");

            }
            finally
            {
                jobNotifications.DeleteFirstJob();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_pagination()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            JobNotifications jobNotifications = homePage.GoToJobs_Settings().GoTo_JOB_NOTIFICATION();
            jobNotifications.PageSize("8");
            var resultPageSize = jobNotifications.GetEmailList().Count();
            jobNotifications.GoToPageTwo();
            Assert.IsTrue((resultPageSize <= 8) && (jobNotifications.GetEmailList().Count() <= 8), "La pagination du 8 ne fonctionne pas.");
            jobNotifications.PageSize("16");
            resultPageSize = jobNotifications.GetEmailList().Count();
            jobNotifications.GoToPageTwo();
            Assert.IsTrue((resultPageSize <= 16) && (jobNotifications.GetEmailList().Count() <= 16), "La pagination du 16 ne fonctionne pas.");
            jobNotifications.PageSize("30");
            resultPageSize = jobNotifications.GetEmailList().Count();
            jobNotifications.GoToNextPage();
            Assert.IsTrue((resultPageSize <= 30) && (jobNotifications.GetEmailList().Count() <= 30), "La pagination du 30 ne fonctionne pas.");
            jobNotifications.PageSize("50");
            resultPageSize = jobNotifications.GetEmailList().Count();
            jobNotifications.GoToNextPage();
            Assert.IsTrue((resultPageSize <= 50) && (jobNotifications.GetEmailList().Count() <= 50), "La pagination du 50 ne fonctionne pas.");
            jobNotifications.PageSize("100");
            if (jobNotifications.GetEmailList().Count() >= 100)
            {
                jobNotifications.GoToNextPage();
            }
            Assert.IsTrue(jobNotifications.GetEmailList().Count() <= 100, "La pagination du 100 ne fonctionne pas.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_AddNew_FileFlowJobValidator()
        {
            
            var homePage = LogInAsAdmin();
            var jobNotifications = homePage.GoToJobs_Settings().GoTo_JOB_NOTIFICATION();
            var fileFlowModal = jobNotifications.AddNewFileFlowJobNotification();
            fileFlowModal.SubmitForm();
            Assert.IsTrue(fileFlowModal.CheckValidator(), "Les validators n'apparaissent pas");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_AddNew_SageJob()
        {
            Random rnd = new Random();
            var email = CUSTOMEREMAIL + rnd.Next().ToString();
            var homePage = LogInAsAdmin();

            var jobsetting = homePage.GoToJobs_Settings();
            var jobnotification = jobsetting.GoTo_JOB_NOTIFICATION();
            try
            {
                var Creatsagejobnotification = jobnotification.AddNewSageJobNotification();
                Creatsagejobnotification.FillField_AddNewsagejobnotification(email);
                jobnotification.Filter(JobNotifications.FilterType.Search, email);
                jobnotification.Filter(JobNotifications.FilterType.NotificationTypes, SAGE);
                Assert.IsTrue(jobnotification.FindRowEmail(email), "La notificaiton n'apparait pas dans la liste");
            }
            finally
            {
                jobnotification.DeleteFirstJob();
            }



        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_AddNew_SageJobValidator()
        {
 
            var homePage = LogInAsAdmin();
            SettingsPage jobsetting = homePage.GoToJobs_Settings();
            JobNotifications jobnotification = jobsetting.GoTo_JOB_NOTIFICATION();
            var Creatsagejobnotification = jobnotification.AddNewSageJobNotification();
            Creatsagejobnotification.FillField_AddNewsagejobnotification("");
            Assert.IsTrue(Creatsagejobnotification.VerifyEmailRequired(), "Les validators n'apparaissent pas");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_filters_NotificationTypes()
        {
            //Prepare
            var type = "Converter";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var jobNotifications = homePage.GoToJobs_Settings().GoTo_JOB_NOTIFICATION();
            jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, "Converter");
            var typelist = jobNotifications.GetTypeList();
            bool typeFound = typelist.Any(item => item.Contains(type));

            //Assert
            Assert.IsTrue(typeFound, $"This type '{type}'n'a pas été trouvé dans la liste.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_AddNew_CustomerJobValidator()
        {
            //Arrange
            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            var jobnotification = jobsetting.GoTo_JOB_NOTIFICATION();
            var customerjobnotificationmodel = jobnotification.AddNewCustomerJobNotification();
            customerjobnotificationmodel.CreateCustomerJobNotificationButton();
            var isEmailValidatorExist = customerjobnotificationmodel.GetEmailRequiredMessage();
            var isCoustomerValidatorExist = customerjobnotificationmodel.GetCustomerRequiredMessage();
            Assert.IsTrue(isEmailValidatorExist, "validator Email n'apparait pas.");
            Assert.IsTrue(isCoustomerValidatorExist, "validator Customer n'apparait pas");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_Converters_AddNew()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string converterType = "JAF_csv";

            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            var settingsTabConvertersPage = jobsetting.GoTo_CONVERTERS();
            try
            {
                var convertersCreateModalPage = settingsTabConvertersPage.AddNewConverter();
                var converterDetailsPage = convertersCreateModalPage.FillFiled_CreateConverters(customer, converterType);
                settingsTabConvertersPage = converterDetailsPage.BackToList();
                settingsTabConvertersPage.ResetFilter();
                settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.Customers, customer);
                settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.ConverterType, converterType);
                var verifyCustomer = settingsTabConvertersPage.GetCustomersFiltred().All(d => d.Equals(customer));
                var verifyConverter = settingsTabConvertersPage.GetConverterTypeFiltred().All(d => d.Equals(converterType));
                var dataCorrect = verifyCustomer && verifyConverter;
                Assert.IsTrue(dataCorrect, "customer et converter type n'apparaitre pas.");
            }
            finally
            {
                settingsTabConvertersPage.deleteFirstConverters();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_GuestMapping()
        {
            var MappingName = "MapTest";

            var homePage = LogInAsAdmin();
            SettingsPage settingsPage = homePage.GoToJobs_Settings();
            GuestMappingsPage guestMappingsPage = settingsPage.GoTo_GUEST_MAPPINGS();
            try 
            {
                guestMappingsPage.EditMapping(MappingName);
                SettingsTabConvertersPage settingsTabConvertersPages = settingsPage.GoTo_CONVERTERS();
                guestMappingsPage = settingsPage.GoTo_GUEST_MAPPINGS();
                var firstMapping = guestMappingsPage.GetFirstMapping();
                Assert.AreEqual(firstMapping, MappingName, MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE);
            }
            finally
            {
                guestMappingsPage.EditMapping("");
                SettingsTabConvertersPage settingsTabConvertersPages = settingsPage.GoTo_CONVERTERS();
                guestMappingsPage = settingsPage.GoTo_GUEST_MAPPINGS();
                var firstMapping = guestMappingsPage.GetFirstMapping();
                Assert.AreEqual(firstMapping, "", MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE);
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_JobNotifications_AddNew_CustomerJob()
        {
            Random rnd = new Random();
            var number = rnd.Next(1, 999).ToString();
            var email = CUSTOMEREMAIL + rnd.Next().ToString();
            string customer = TestContext.Properties["CustomerJob"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();
            var jobsetting = homePage.GoToJobs_Settings();
            var jobnotification = jobsetting.GoTo_JOB_NOTIFICATION();
            var customerJobNotificationModel = jobnotification.AddNewCustomerJobNotification();
            customerJobNotificationModel.FillFieldAddNewCustomerJobNotification(email, customer);
            jobnotification.Filter(JobNotifications.FilterType.Search, email);
            var enteredData = jobnotification.GetJobNotifLine();
            bool emailAndCustomerFound = enteredData.Any(item => item.Contains(email) && item.Contains(customer));
            Assert.IsTrue(emailAndCustomerFound, $"L'email '{email}' et le customer '{customer}' n'ont pas été trouvés dans la liste. Le New CUSTOMER JOB n'a pas été ajouté correctement.");
        }

        [TestMethod]
        [Priority(3)]
        [Timeout(_timeout)]
        public void JB_SETT_JobNotifications_CreateJobNotificationForTests()
        {
            string fiscalEntities = "NEWREST ESPAÑA S.L.";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var jobSettings = homePage.GoToJobs_Settings();
            var jobNotifications = jobSettings.GoTo_JOB_NOTIFICATION();
            jobNotifications.Filter(JobNotifications.FilterType.Search, CUSTOMEREMAIL);
            jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
            jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, NOTIFTYPECUSTOMER);
            if (jobNotifications.CountJobs() == 0)
            {
                var customerModal = jobNotifications.AddNewCustomerJobNotification();
                customerModal.FillFieldAddNewCustomerJobNotification(CUSTOMEREMAIL, CUSTOMER);
                jobNotifications.ResetFilter();
                jobSettings.GoTo_JOB_NOTIFICATION();
                jobNotifications.Filter(JobNotifications.FilterType.Search, CUSTOMEREMAIL);
                jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
                jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, NOTIFTYPECUSTOMER);
                Assert.AreEqual(1, jobNotifications.CountJobs(), "email et customer n'apparaitre pas.");
            }
            jobNotifications.ResetFilter();
            jobSettings.GoTo_JOB_NOTIFICATION();
            jobNotifications.Filter(JobNotifications.FilterType.Search, CONVERTEREMAIL);
            jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
            jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, CONVERTER);

            if (jobNotifications.CountJobs() == 0)
            {
                var ConvertersPage = jobSettings.GoTo_CONVERTERS();
                ConvertersPage.Filter(SettingsTabConvertersPage.FilterType.Customers, CUSTOMER);
                ConvertersPage.Filter(SettingsTabConvertersPage.FilterType.ConverterType, CONVERTERTYPE);
                if (ConvertersPage.GetCustomersFiltred().Count == 0)
                {
                    var convertersCreateModalPage = ConvertersPage.AddNewConverter();
                    var converterDetailsPage = convertersCreateModalPage.FillFiled_CreateConverters(CUSTOMER, CONVERTERTYPE);
                    ConvertersPage = converterDetailsPage.BackToList();
                    jobNotifications = ConvertersPage.GoTo_JOB_NOTIFICATION();
                    var ConverterModal = jobNotifications.AddNewConverterJobNotification();
                    ConverterModal.FillField_AddNewCONVERTERSJOBNOTIFICATION(CONVERTEREMAIL, CUSTOMER, CONVERTERTYPE);
                    jobNotifications.ResetFilter();
                    jobSettings.GoTo_JOB_NOTIFICATION();
                    jobNotifications.Filter(JobNotifications.FilterType.Search, CONVERTEREMAIL);
                    jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
                    jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, CONVERTER);
                    Assert.AreEqual(1, jobNotifications.CountJobs(), "email et customer n'apparaitre pas.");
                }
                else
                {
                    jobSettings.GoTo_JOB_NOTIFICATION();
                    var ConverterModal = jobNotifications.AddNewConverterJobNotification();
                    ConverterModal.FillField_AddNewCONVERTERSJOBNOTIFICATION(CONVERTEREMAIL, CUSTOMER, CONVERTERTYPE);
                    jobNotifications.ResetFilter();
                    jobSettings.GoTo_JOB_NOTIFICATION();
                    jobNotifications.Filter(JobNotifications.FilterType.Search, CONVERTEREMAIL);
                    jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
                    jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, CONVERTER);
                    Assert.AreEqual(1, jobNotifications.CountJobs(), "email et customer n'apparaitre pas.");
                }

            }
            jobNotifications.ResetFilter();
            jobSettings.GoTo_JOB_NOTIFICATION();
            jobNotifications.Filter(JobNotifications.FilterType.Search, FILEFLOWEMAIL);
            jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
            jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, FILEFLOW);

            if (jobNotifications.CountJobs() == 0)
            {
                var ScheduledPage = homePage.GoToJobs_ScheduledJobs();
                var FileFlowPage = ScheduledPage.Nav_To_File_Flows();
                FileFlowPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.Customers, CUSTOMER);
                FileFlowPage.Filter(ScheduledJobsTabFileFlowsPage.FilterType.Converter, CONVERTERTYPE);
                if (FileFlowPage.GetTotalList() == 0)
                {
                    var FileFlowsCreateModal = FileFlowPage.OpenModalCreateFileFLows();
                    FileFlowsCreateModal.FillFiled_CreateFileFlows(FILEFLOWEMAIL, CUSTOMER, true, CONVERTERTYPE);
                    jobNotifications = ScheduledPage.GoToJobs_Settings().GoTo_JOB_NOTIFICATION();
                    var FileFlowModal = jobNotifications.AddNewFileFlowJobNotification();
                    FileFlowModal.FillFieldAddNewFileFlowJobNotification(FILEFLOWEMAIL, CUSTOMER, FILE_FLOW_NAME);
                    jobNotifications.ResetFilter();
                    jobSettings.GoTo_JOB_NOTIFICATION();
                    jobNotifications.Filter(JobNotifications.FilterType.Search, FILEFLOWEMAIL);
                    jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
                    jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, FILEFLOW);
                    Assert.AreEqual(1, jobNotifications.CountJobs(), "email et customer n'apparaitre pas.");
                }
                else
                {
                    jobNotifications = ScheduledPage.GoToJobs_Settings().GoTo_JOB_NOTIFICATION();
                    var FileFlowModal = jobNotifications.AddNewFileFlowJobNotification();
                    FileFlowModal.FillFieldAddNewFileFlowJobNotification(FILEFLOWEMAIL, CUSTOMER, FILE_FLOW_NAME);
                    jobNotifications.ResetFilter();
                    jobSettings.GoTo_JOB_NOTIFICATION();
                    jobNotifications.Filter(JobNotifications.FilterType.Search, FILEFLOWEMAIL);
                    jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
                    jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, FILEFLOW);
                    Assert.AreEqual(1, jobNotifications.CountJobs(), "email et customer n'apparaitre pas.");
                }

            }
            jobNotifications.ResetFilter();
            jobSettings.GoTo_JOB_NOTIFICATION();
            jobNotifications.Filter(JobNotifications.FilterType.Search, SAGEEMAIL);
            jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
            jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, SAGE);
            if (jobNotifications.CountJobs() == 0)
            {
                var ScheduledPage = homePage.GoToJobs_ScheduledJobs();
                var scheduledJobsCegidPage = ScheduledPage.Nav_To_Cegid();
                scheduledJobsCegidPage.ResetFilter();
                scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.SearchByName, CEGIDSCHEDULED);
                scheduledJobsCegidPage.Filter(ScheduledJobsCegidPage.FilterType.ShowActiveOnly, true);
                if (!scheduledJobsCegidPage.GetAllNamesResultPaged().Contains(CEGIDSCHEDULED))
                {
                    CegidScheduledJobModal newCegidScheduledModal = scheduledJobsCegidPage.NewCegidScheduledModal();
                    newCegidScheduledModal.FillFieldsModalCegidScheduledJob(CEGIDSCHEDULED, fiscalEntities);
                    jobNotifications = ScheduledPage.GoToJobs_Settings().GoTo_JOB_NOTIFICATION();
                    var SageModal = jobNotifications.AddNewSageJobNotification();
                    SageModal.FillField_AddNewsagejobnotification(SAGEEMAIL, CEGIDSCHEDULED);
                    jobNotifications.ResetFilter();
                    jobSettings.GoTo_JOB_NOTIFICATION();
                    jobNotifications.Filter(JobNotifications.FilterType.Search, SAGEEMAIL);
                    jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
                    jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, SAGE);
                    Assert.AreEqual(1, jobNotifications.CountJobs(), "email et customer n'apparaitre pas.");
                }
                else
                {
                    jobNotifications = ScheduledPage.GoToJobs_Settings().GoTo_JOB_NOTIFICATION();

                    var SageModal = jobNotifications.AddNewSageJobNotification();
                    SageModal.FillField_AddNewsagejobnotification(SAGEEMAIL, CEGIDSCHEDULED);
                    jobNotifications.ResetFilter();
                    jobSettings.GoTo_JOB_NOTIFICATION();
                    jobNotifications.Filter(JobNotifications.FilterType.Search, SAGEEMAIL);
                    jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
                    jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, SAGE);
                    Assert.AreEqual(1, jobNotifications.CountJobs(), "email et customer n'apparaitre pas.");
                }
            }
            jobNotifications.ResetFilter();
            jobSettings.GoTo_JOB_NOTIFICATION();
            jobNotifications.Filter(JobNotifications.FilterType.Search, SPECIFICEMAIL);
            jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
            jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, SPECIFIC);
            if (jobNotifications.CountJobs() == 0)
            {
                var ScheduledPage = homePage.GoToJobs_ScheduledJobs();
                var scheduledSpecificJobsPage = ScheduledPage.Nav_To_Specific();
                scheduledSpecificJobsPage.Filter(ScheduledSpecificPage.FilterType.SearchByName, SPECIFICSCHEDULE);
                scheduledSpecificJobsPage.Filter(ScheduledSpecificPage.FilterType.ShowActiveOnly, true);
                if (scheduledSpecificJobsPage.CountSpecificJobs() == 0)
                {
                    var newSpecificJobdModal = scheduledSpecificJobsPage.OpenModalCreateSpecificJob();
                    newSpecificJobdModal.FillFiled_CreateSpecific(SPECIFICSCHEDULE, true);
                    jobNotifications = ScheduledPage.GoToJobs_Settings().GoTo_JOB_NOTIFICATION();
                    var SpecificModal = jobNotifications.AddNewSpecificJobNotification();
                    SpecificModal.FillField_AddNewspecificjobnotification(SPECIFICEMAIL, SPECIFICSCHEDULE);
                    jobNotifications.ResetFilter();
                    jobSettings.GoTo_JOB_NOTIFICATION();
                    jobNotifications.Filter(JobNotifications.FilterType.Search, SPECIFICEMAIL);
                    jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
                    jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, SPECIFIC);
                    Assert.AreEqual(1, jobNotifications.CountJobs(), "email et customer n'apparaitre pas.");
                }
                else
                {
                    jobNotifications = ScheduledPage.GoToJobs_Settings().GoTo_JOB_NOTIFICATION();
                    var SpecificModal = jobNotifications.AddNewSpecificJobNotification();
                    SpecificModal.FillField_AddNewspecificjobnotification(SPECIFICEMAIL, SPECIFICSCHEDULE);
                    jobNotifications.ResetFilter();
                    jobSettings.GoTo_JOB_NOTIFICATION();
                    jobNotifications.Filter(JobNotifications.FilterType.Search, SPECIFICEMAIL);
                    jobNotifications.Filter(JobNotifications.FilterType.Customer, CUSTOMER);
                    jobNotifications.Filter(JobNotifications.FilterType.NotificationTypes, SPECIFIC);
                    Assert.AreEqual(1, jobNotifications.CountJobs(), "email et customer n'apparaitre pas.");
                }

            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_NewConverter()
        {
            var homePage = LogInAsAdmin();
            var jobSettingPage = homePage.GoToJobs_Settings();
            var settingsTabConvertersPage = jobSettingPage.GoTo_CONVERTERS();
            var convertersCreateModalPage = settingsTabConvertersPage.AddNewConverter();
            bool areAllAligned = convertersCreateModalPage.VerifyThreeLabelsAndInputsInSameRow();
            Assert.IsTrue(areAllAligned, "Les champs et les titres ne sont pas tous alignés");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_SETT_Converters_AddNewValidator()
        {
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string converterType = "NoConverter";

            // Arrange
            var homePage = LogInAsAdmin();
            var jobSettingPage = homePage.GoToJobs_Settings();
            var settingsTabConvertersPage = jobSettingPage.GoTo_CONVERTERS();
            var convertersCreateModalPage = settingsTabConvertersPage.AddNewConverter();
            //simulate empty file extension and no converter
            convertersCreateModalPage.FillFiled_CreateConverters(customer, converterType, "");
            var verifyValidator = convertersCreateModalPage.VerifyCreateValidators();
            Assert.IsTrue(verifyValidator, "Les validateurs n'apparaissent pas tous !");
        }

        [TestMethod]
        public void JB_SETT_ConvertersTP06_IncludedPacketPolicyGroups()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerJob"].ToString();
            string converterType = "AV_TP06_XML";
            string value = "AB(35484,35485,61133,65677,68345,68346,68347,72053,72055,77296,77298,85087);AM(30567,35754);BV(4717,62811,62812,62813,75205,1776,1787,1790,3558,4715,14996,34732,51792,51793,52921,60165,62575,68177,68178,72044,85811);CP;DH(59153,60813,86818,86819);EQ(2495,19958,19960,20049,20090,25170,38395,52750,52751,52752,52753,55006,55007,59001,62614,62706,62707,65721,68115,68170,1831,41739,4717,58339,58340,58344,62613,62615,64209,64342,71118,82697,827002);FD(1831,23043,23755,51793,51821,62706,74723,82008,82019,85656,86363,86374,86441,86444,86868);LB(2521,35483,35484,35485,56993,56994);MU(21054,21097,21257,21069);OR(56164,68115,86999,35532,54423,54424);SC;SG;SO(35485,57463,68145,68146,68149,68153,86119,86122);W1;W2";
            string IncludedPacket = "IncludedPacketPolicyGroups"; 
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var jobsetting = homePage.GoToJobs_Settings();
            var settingsTabConvertersPage = jobsetting.GoTo_CONVERTERS();
            try
            {
                var convertersCreateModalPage = settingsTabConvertersPage.AddNewConverter();
                var converterDetailsPage = convertersCreateModalPage.FillFiled_CreateConverters(customer, converterType);
                settingsTabConvertersPage = converterDetailsPage.BackToList();
                settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.Customers, customer);
                settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.ConverterType, converterType);

                var ConverterLignePage = settingsTabConvertersPage.SelectFirstItem();
                ConverterLignePage.GoToSettingsConverter();

                var SettingCreateModalPage = ConverterLignePage.ClickNewSettingCreatePage();
                converterDetailsPage = SettingCreateModalPage.FillField_CreatNewSettingConverter(value, IncludedPacket);

                bool isCreated = converterDetailsPage.IscreatedOKSettingsConverter(value);

                //Le test réussit si la création a été effectuée, sinon le test échoue
                Assert.IsTrue(isCreated, "Pas d'Erreur le Converter doit prendre en compte tous les codes services rentrés par l'utilisateur.");
                settingsTabConvertersPage = converterDetailsPage.BackToList();

            }
            finally
            {
                settingsTabConvertersPage.ResetFilter(); 
                settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.Customers, customer);
                settingsTabConvertersPage.Filter(SettingsTabConvertersPage.FilterType.ConverterType, converterType);
                settingsTabConvertersPage.deleteFirstConverters();
            }
        }

    }
}
