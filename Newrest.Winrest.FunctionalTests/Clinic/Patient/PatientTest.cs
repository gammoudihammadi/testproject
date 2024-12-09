using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using static Newrest.Winrest.FunctionalTests.PageObjects.Clinic.Patient.PatientsPage;

namespace Newrest.Winrest.FunctionalTests.Clinic
{
    [TestClass]
    public class PatientTest : TestBase
    {
        private const int _timeout = 600000;

        // données tests
        readonly static string today = DateUtils.Now.ToString("dd/MM/yyyy");

        // Ajouter l'offre
        [TestMethod]
        [Priority(1)]
        [Timeout(_timeout)]
        public void CL_PAT_SetClinicParameters()
        {
            //Prepare
            string siteName = TestContext.Properties["ClinicNephrocare"].ToString();
            string siteCode = TestContext.Properties["ClinicSiteCode"].ToString();

            var adminName = TestContext.Properties["Admin_UserName"].ToString();
            string userName = adminName.Substring(0, adminName.IndexOf("@"));

            string offerName = TestContext.Properties["ClinicOfferStandard"].ToString();
            string textureName = TestContext.Properties["ClinicTextureNormal"].ToString();
            string customerTypeClinics = TestContext.Properties["ClinicCustomerType"].ToString();
            string clinicGuest = TestContext.Properties["ClinicGuest"].ToString();
            string guestTypeType = TestContext.Properties["ClinicGuestTypeTypeRefClinic"].ToString();

            string mealVariantBreakfast = TestContext.Properties["ClinicMealBreakfast"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Create a new Site
            var parametersSites = homePage.GoToParameters_Sites();
            parametersSites.Filter(ParametersSites.FilterType.SearchSite, siteName);
            bool isSite = parametersSites.isSiteExists(siteName);

            if (!isSite)
            {
                var sitesModalPage = parametersSites.ClickOnNewSite();
                sitesModalPage.FillPrincipalField_CreationNewSite(siteName, "-", "-", "-", siteCode);
            }
            parametersSites.Filter(ParametersSites.FilterType.SearchSite, siteName);
            string id = parametersSites.CollectNewSiteID();
            // Affect it to the user
            var parametersUser = homePage.GoToParameters_User();
            parametersUser.SearchAndSelectUser(userName);
            parametersUser.ClickOnAffectedSite();
            parametersUser.GiveSiteRightsToUser(id, true, siteName);

            //Parameters Clinic
            var parametersClinic = homePage.GoToParameters_Clinic();
            parametersClinic.ClickToOffer();
            bool isOffer = parametersClinic.isOfferExist(offerName);
            if (!isOffer)
            {
                parametersClinic.AddNewOffer(offerName, "1");
            }

            parametersClinic.ClickToPackage();
            bool isPackage = parametersClinic.isPackageExist(offerName);
            if (!isPackage)
            {
                parametersClinic.AddNewPackage(offerName, "1");
            }

            parametersClinic.ClickToTexture();
            bool isTexture = parametersClinic.isPackageExist(textureName);
            if (!isTexture)
            {
                parametersClinic.AddNewTexture(textureName, "1");
            }

            //Parameters Customer
            var parametersCustomer = homePage.GoToParameters_CustomerPage();
            bool isCustomerType = parametersCustomer.isTypeOfCustomerExist(customerTypeClinics);
            if (!isCustomerType)
            {
                parametersCustomer.AddNewTypeOfCustomer(customerTypeClinics);
            }

            //Parameters Production - Guest
            var parametersProduction = homePage.GoToParameters_ProductionPage();
            parametersProduction.GoToTab_Guest();
            if (!parametersProduction.IsGuestPresent(clinicGuest))
            {
                parametersProduction.AddNewGuest(clinicGuest, false, "500");
                parametersProduction.EditGuest(clinicGuest);
                parametersProduction.AddIsClinicToGuest();
                parametersProduction.EditGuest(clinicGuest);
                parametersProduction.AddGuestTypeTypeToGuest(guestTypeType);
            }

            //Parameters Production - MEAL
            parametersProduction.GoToTab_Meal();
            if (!parametersProduction.IsMealPresent(mealVariantBreakfast))
            {
                parametersProduction.CreateNewMeal(mealVariantBreakfast, mealVariantBreakfast, "9", "06:00");
            }

            // Parameters Production - Variant
            parametersProduction.GoToTab_Variant();
            parametersProduction.FilterSite(siteName);
            parametersProduction.FilterGuestType(clinicGuest);
            parametersProduction.FilterMealType(mealVariantBreakfast);

            if (!parametersProduction.IsVariantPresent(clinicGuest, mealVariantBreakfast))
            {
                parametersProduction.AddNewVariant(clinicGuest, mealVariantBreakfast, siteName);
            }
        }

        //Création d'un nouveau site pour les tests Clinic
        [TestMethod]
        [Priority(2)]
        [Timeout(_timeout)]
        public void CL_PAT_Create_Clinic_Tests_Data()
        {
            //Prepare
            string siteName = TestContext.Properties["ClinicNephrocare"].ToString();

            string customerName = TestContext.Properties["ClinicCustomerName"].ToString();
            string customerIcao = TestContext.Properties["ClinicCustomerIcao"].ToString();
            string customerTypeClinics = TestContext.Properties["ClinicCustomerType"].ToString();
            string deliveryName = TestContext.Properties["ClinicDelivery"].ToString();

            //data for more tests
            string mealVariantBreakfast = TestContext.Properties["ClinicGuest"].ToString() + " - " + TestContext.Properties["ClinicMealBreakfast"].ToString();

            string menuName = "Menu Diabete Clinic Test";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Create Customer Clinique
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerIcao, customerTypeClinics);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerName);
                Assert.IsTrue(customersPage.GetFirstCustomerName().Contains(customerName), "Le customer n'a pas été créé.");
            }

            // Create a new Delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, siteName, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryName);
                Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(deliveryName), "La delivery n'a pas été crée");
            }

            //Create menu is Clinic
            /*var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            if (menusPage.CheckTotalNumber() == 0)
            {
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), siteName, mealVariantBreakfast);
            var menusGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
            menusGeneralInformationPage.SetClinic(true);
            }*/
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CL_PAT_Create_New_Patient()
        {
            // Prepare
            string rndNb = new Random().Next(0, 1000000).ToString("D6");
            string patientFirstName = "Patient " + rndNb;
            string patientLastName = "Test " + DateUtils.Now.ToString("dd/MM/yyyy");
            string deliveryName = TestContext.Properties["ClinicDelivery"].ToString();
            string siteCode = $"{TestContext.Properties["ClinicSiteCode"].ToString()} - {TestContext.Properties["ClinicNephrocare"].ToString()}";


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var patientsPage = homePage.GoToClinic_PatientPage();
            var patientCreateModalPage = patientsPage.PatientCreatePage();
            patientCreateModalPage.FillFields_CreatePatientModalPage(patientFirstName, patientLastName, rndNb, siteCode, deliveryName);

            patientsPage.ResetFilters();
            patientsPage.Filter(FilterType.Search, patientFirstName);

            //Asserts
            var firstNameDisplay = patientsPage.GetFirstPatientFirstName();
            Assert.AreEqual(patientFirstName, firstNameDisplay, "Le patient n'a pas été créé.");
            Assert.AreEqual(patientLastName, patientsPage.GetFirstPatientLastName(), "Le 'Last name' du patient n'est pas le bon.");
            Assert.AreEqual(rndNb, patientsPage.GetFirstPatientIpp(), "L''IPP' du patient n'est pas le bon.");
            Assert.AreEqual(rndNb, patientsPage.GetFirstPatientVisitNumber(), "Le 'Visit number' du patient n'est pas le bon.");
            Assert.AreEqual(deliveryName, patientsPage.GetFirstPatientFloor(), "Le 'Floor' du patient n'est pas le bon.");
            Assert.AreEqual(rndNb + " / " + rndNb, patientsPage.GetFirstPatientRoomBed(), "La 'Room / Bed' du patient n'est pas correcte.");
            Assert.AreEqual(today, patientsPage.GetFirstPatientStartDate(), "La 'Start Date' du patient n'est pas correcte.");
            Assert.AreEqual(DateUtils.Now.AddDays(7).ToString("dd/MM/yyyy"), patientsPage.GetFirstPatientEndDate(), "La 'End Date' du patient n'est pas correcte.");
            Assert.IsTrue(patientsPage.GetFirstPatientDietMonitoring().Contains("notConcerned"), "Le 'Diet monitoring' du patient ne correspond pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CL_PAT_Delete_Patient()
        {
            string rndNb = new Random().Next(0, 1000000).ToString("D6");
            string patientFirstName = $"Patient {rndNb}";
            string patientLastName = $"Test {DateUtils.Now.ToString("dd/MM/yyyy")}";
            string deliveryName = TestContext.Properties["ClinicDelivery"].ToString();
            string siteCode = $"{TestContext.Properties["ClinicSiteCode"].ToString()} - {TestContext.Properties["ClinicNephrocare"].ToString()}";
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            PageObjects.Clinic.Patient.PatientsPage patientsPage = homePage.GoToClinic_PatientPage();
            PageObjects.Clinic.Patient.PatientCreateModalPage patientCreateModalPage = patientsPage.PatientCreatePage();
            patientCreateModalPage.FillFields_CreatePatientModalPage(patientFirstName, patientLastName, rndNb, siteCode, deliveryName);
            patientsPage.ResetFilters();
            patientsPage.Filter(FilterType.Search, patientFirstName);
            string firstPatientFirstName = patientsPage.GetFirstPatientFirstName();
            Assert.AreEqual(patientFirstName, firstPatientFirstName, "Le patient n'a pas été créé.");
            patientsPage.DeletePatient(patientFirstName);
            patientsPage.Filter(FilterType.Search, patientFirstName);
            bool isPatientDisplayed = patientsPage.IsPatientDisplayed();
            Assert.IsFalse(isPatientDisplayed, "Le patient n'a pas été supprimé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CL_PAT_Filter_Date()
        {
            // Prepare
            string rndNb = new Random().Next(0, 1000000).ToString("D6");
            string patientFirstName = "Patient " + rndNb;
            string patientLastName = "Test " + DateUtils.Now.ToString("dd/MM/yyyy");
            string deliveryName = TestContext.Properties["ClinicDelivery"].ToString();
            string siteCode = $"{TestContext.Properties["ClinicSiteCode"].ToString()} - {TestContext.Properties["ClinicNephrocare"].ToString()}";
            DateTime fromDate = DateUtils.Now.AddDays(-1);
            DateTime toDate = DateUtils.Now.AddDays(2);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            //1. Filtre
            var patientsPage = homePage.GoToClinic_PatientPage();

            //2. Créer un patient si pas assez
            if (patientsPage.CheckTotalNumber() < 20)
            {
                var patientCreateModalPage = patientsPage.PatientCreatePage();
                patientCreateModalPage.FillFields_CreatePatientModalPage(patientFirstName, patientLastName, rndNb, siteCode, deliveryName);
            }

            patientsPage.ResetFilters();
            patientsPage.Filter(FilterType.DateFrom, fromDate);
            patientsPage.Filter(FilterType.DateTo, toDate);

            if (!patientsPage.isPageSizeEqualsTo100())
            {
                patientsPage.PageSize("8");
                patientsPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(patientsPage.IsDateRespected(fromDate, toDate), String.Format(MessageErreur.FILTRE_ERRONE, "From and To"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CL_PAT_Filter_MonitoringDiet_NotConcerned()
        {
            // Prepare
            string rndNb = new Random().Next(0, 1000000).ToString("D6");
            string patientFirstName = "Patient " + rndNb;
            string patientLastName = "Test " + DateUtils.Now.ToString("dd/MM/yyyy");
            string deliveryName = TestContext.Properties["ClinicDelivery"].ToString();
            string siteCode = TestContext.Properties["ClinicSiteCode"].ToString() + " - " + TestContext.Properties["ClinicNephrocare"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            //1. Filtre
            var patientsPage = homePage.GoToClinic_PatientPage();

            //2. Créer un patient si pas assez
            if (patientsPage.CheckTotalNumber() < 20)
            {
                var patientCreateModalPage = patientsPage.PatientCreatePage();
                patientCreateModalPage.FillFields_CreatePatientModalPage(patientFirstName, patientLastName, rndNb, siteCode, deliveryName);
            }

            patientsPage.ResetFilters();
            patientsPage.ClickDietMonitoring();
            patientsPage.Filter(FilterType.DietMonitoringNotConcerned, true);

            if (!patientsPage.isPageSizeEqualsTo100())
            {
                patientsPage.PageSize("8");
                patientsPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(patientsPage.CheckDietMonitoring("icon-FourknKnives notConcerned"), String.Format(MessageErreur.FILTRE_ERRONE, "'Diet Monitoring Not Concerned'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CL_PAT_Filter_MonitoringDiet_Done()
        {
            // Prepare
            string rndNb = new Random().Next(0, 1000000).ToString("D6");
            string patientFirstName = "Patient " + rndNb;
            string patientLastName = "Test " + DateUtils.Now.ToString("dd/MM/yyyy");
            string deliveryName = TestContext.Properties["ClinicDelivery"].ToString();
            string siteCode = TestContext.Properties["ClinicSiteCode"].ToString() + " - " + TestContext.Properties["ClinicNephrocare"].ToString();

            // Arrange
            HomePage homePage= LogInAsAdmin();
            
            // Act
            //1. Filtre
            var patientsPage = homePage.GoToClinic_PatientPage();

            //2. Créer un patient diet done si pas assez
            patientsPage.ResetFilters();
            patientsPage.ClickDietMonitoring();
            patientsPage.Filter(FilterType.DietMonitoringDone, true);
            if (patientsPage.CheckTotalNumber() < 20)
            {
                var patientCreateModalPage = patientsPage.PatientCreatePage();
                patientCreateModalPage.FillFields_CreatePatientModalPage(patientFirstName, patientLastName, rndNb, siteCode, deliveryName, "Monitoring done");
                patientsPage.ResetFilters();
                patientsPage.Filter(FilterType.Search, patientFirstName);

                //Asserts
                Assert.AreEqual(patientFirstName, patientsPage.GetFirstPatientFirstName(), "Le patient avec la diet monitoring 'Done' n'a pas été créé.");
                patientsPage.ResetFilters();
                patientsPage.ClickDietMonitoring();
                patientsPage.Filter(FilterType.DietMonitoringDone, true);
            }

            patientsPage.ResetFilters();
            patientsPage.ClickDietMonitoring();
            patientsPage.Filter(FilterType.DietMonitoringDone, true);

            if (!patientsPage.isPageSizeEqualsTo100())
            {
                patientsPage.PageSize("8");
                patientsPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(patientsPage.CheckDietMonitoring("icon-FourknKnives valid"), String.Format(MessageErreur.FILTRE_ERRONE, "'Diet Monitoring Done'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CL_PAT_Filter_MonitoringDiet_NotDone()
        {
            // Prepare
            string rndNb = new Random().Next(0, 1000000).ToString("D6");
            string patientFirstName = "Patient " + rndNb;
            string patientLastName = "Test " + DateUtils.Now.ToString("dd/MM/yyyy");
            string deliveryName = TestContext.Properties["ClinicDelivery"].ToString();
            string siteCode = TestContext.Properties["ClinicSiteCode"].ToString() + " - " + TestContext.Properties["ClinicNephrocare"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            //1. Filtre
            var patientsPage = homePage.GoToClinic_PatientPage();

            //2. Créer un patient si pas assez
            patientsPage.ResetFilters();
            patientsPage.ClickDietMonitoring();
            patientsPage.Filter(FilterType.DietMonitoringNotDone, true);
            if (patientsPage.CheckTotalNumber() < 20)
            {
                var patientCreateModalPage = patientsPage.PatientCreatePage();
                patientCreateModalPage.FillFields_CreatePatientModalPage(patientFirstName, patientLastName, rndNb, siteCode, deliveryName, "Monitoring not done");
                patientsPage.ResetFilters();
                patientsPage.Filter(FilterType.Search, patientFirstName);

                //Asserts
                Assert.AreEqual(patientFirstName, patientsPage.GetFirstPatientFirstName(), "Le patient avec la diet monitoring 'Done' n'a pas été créé.");
                patientsPage.ResetFilters();
                patientsPage.ClickDietMonitoring();
                patientsPage.Filter(FilterType.DietMonitoringNotDone, true);
            }

            if (!patientsPage.isPageSizeEqualsTo100())
            {
                patientsPage.PageSize("8");
                patientsPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(patientsPage.CheckDietMonitoring("icon-FourknKnives warning"), String.Format(MessageErreur.FILTRE_ERRONE, "'Diet Monitoring Done'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CL_PAT_Filter_MonitoringDiet_All()
        {
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            //1. Filtre
            var patientsPage = homePage.GoToClinic_PatientPage();

            patientsPage.ResetFilters();
            patientsPage.ClickDietMonitoring();
            patientsPage.Filter(FilterType.DietMonitoringDone, true);
            var doneNb = patientsPage.CheckTotalNumber();

            patientsPage.Filter(FilterType.DietMonitoringNotDone, true);
            var notDoneNb = patientsPage.CheckTotalNumber();

            patientsPage.Filter(FilterType.DietMonitoringNotConcerned, true);
            var notConcernedNb = patientsPage.CheckTotalNumber();

            patientsPage.Filter(FilterType.ShowAll, true);
            var totalNb = patientsPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(doneNb + notDoneNb + notConcernedNb, totalNb, String.Format(MessageErreur.FILTRE_ERRONE, "Show all"));
        }
    }
}
