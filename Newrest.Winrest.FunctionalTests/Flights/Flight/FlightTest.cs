using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using Limilabs.Client.IMAP;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.TimeBlock;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Trolleys;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Flights;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.User;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Threading;
using System.Web.Routing;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;
using static Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight.FlightPage;
using Page = UglyToad.PdfPig.Content.Page;


namespace Newrest.Winrest.FunctionalTests.Flights
{
  
    [TestClass]
    public class FlightTests : TestBase
    {
        string flightNumber = $"Flight 1 - {DateTime.Today.ToShortDateString()} {new Random().Next(100, 999)}";
        string serviceName = "Service-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();
        [TestInitialize]
        public override void TestInitialize()
        {

            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(FL_FLIG_ViewFlightIndex_FilterEquipmenttype_Unknown):
                    string aircraft_number = Test_Initialize_Creat_AirCraft();
                    Test_Initialize_Creat_flight(aircraft_number);
                    break;
                case nameof(FL_FLIG_ViewFlightIndex_FilterCustomers):
                    Test_Initialize_Creat_flight();
                    break;

                case nameof(FL_FLIG_ViewFlight_FilterHidecancelledflights):
                    Test_Initialize_Creat_Cancelled_flight();
                    break;
                case nameof(FL_FLIG_ViewFlightIndex_FilterDocknumber):
                    Test_Initialize_Creat_flight();
                    break;
                case nameof(FL_FLIG_ViewFlightIndex_FilterDriver):
                    Test_Initialize_Creat_flight();
                    break;
                case nameof(FL_FLIG_Calendar_FilterSite):
                    Test_Initialize__Add_Multiple_Flights();
                    break;
                case nameof(FL_FLIG_VueCalendarUnfoldAll):
                    Test_Initialize_Customer_Price_service();
                    Test_Initialize_Creat_flight();
                    break;
               
                default:
                    break;
            }
        }

        // Impossible de évaluer l’expression, car un frame natif se trouve en haut de la pile des appels.
        private const int _timeout = 60 * 10 * 1000;
        private readonly string serviceNameToday = "Service-" + DateUtils.Now.ToString("dd/MM/yyyy");

        #region // Test Init func 
        public void Test_Initialize_VerifyServicesForFlight()
        {
            // Vérifie l'existence d'un service et le crée s'il n'existe pas

            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string serviceCode = "12345";
            string flightType = "International";

            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);

            string customer = customerLp.Substring(0, customerLp.IndexOf("("));

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var tabletPage = homePage.GoToParametres_Tablet();
            tabletPage.ClickFlightTypeTab();
            //check if flight type international exist
            var exist = tabletPage.VerifyInternationalExistOnSite(flightType, site);
            // if no create the flight type international
            if (!exist)
            {
                tabletPage.CreateNewFlightType(flightType, "Orange", site);
            }
            // else set the color gray
            tabletPage.EditFlightTypeColor(flightType, "Orange", site);

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                // Si aucun service n'est disponible, en créer un nouveau

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(site, customer, fromDate, toDate);
                servicePage = pricePage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceName), $"Le service {serviceName} n'existe pas.");
        }
        public void Test_Initialize_VerifyServicesForFlightBOB()
        {
            // Prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["ServiceBob"].ToString();

            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(30);


            string customer = customerLp.Substring(0, customerLp.IndexOf(" ("));

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, "BUY ON BOARD");
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(customer, site, fromDate, toDate);
                servicePage = pricePage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceName), $"Le service {serviceName} n'existe pas.");
        }

        public void Test_Initialize_Create_Lp_CardBob()
        {
            // Prepare
            string name = TestContext.Properties["LpCartNamebob"].ToString();
            string code = TestContext.Properties["LpCartCodebob"].ToString();

            string nameBis = TestContext.Properties["LpCartNameBis"].ToString();
            string codeBis = TestContext.Properties["LpCartCodeBis"].ToString();

            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string aircraft = "AB310";
            DateTime from = DateUtils.Now.AddDays(-1);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Bob comment";

            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();

            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            lpCartPage.Filter(LpCartPage.FilterType.StartDate, DateTime.Today.AddMonths(-1).AddDays(-11));
            lpCartPage.Filter(LpCartPage.FilterType.EndDate, DateTime.Today.AddMonths(2));

            if (lpCartPage.CheckTotalNumber() == 0)
            {
                // Create
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);

                //Acces Onglet Cart                
                LpCartDetailPage.ClickAddtrolley();
                LpCartDetailPage.AddTrolley(trolleyName);

                //Create LpCart Scheme
                var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");
                LpCartDetailPage.BackToList();
            }
            else
            {
                var lpCartDetailPage = lpCartPage.ClickFirstLpCart();
                // On repousse la date du Lpcart
                var infomationPage = lpCartDetailPage.LpCartGeneralInformationPage();
                infomationPage.UpdateToDateTo();
                lpCartDetailPage.BackToList();
            }
            // Assert
            bool isExistFirstNewLpCart = lpCartPage.CheckTotalNumber() >= 1;
            Assert.IsTrue(isExistFirstNewLpCart, "Le lpCart n'est pas crée");

            lpCartPage.Filter(LpCartPage.FilterType.Search, nameBis);

            if (lpCartPage.CheckTotalNumber() == 0)
            {
                //Create second LpCart
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(codeBis, nameBis, site, customer, aircraft, from, to, comment);
                LpCartDetailPage.BackToList();
            }
            else
            {
                var lpCartDetailPage = lpCartPage.ClickFirstLpCart();
                // On repousse la date du Lpcart
                var infomationPage = lpCartDetailPage.LpCartGeneralInformationPage();
                infomationPage.UpdateToDateTo();
                lpCartDetailPage.BackToList();
            }
            // Assert
            bool isExistSecondLpCart = lpCartPage.CheckTotalNumber() > 0;
            Assert.IsTrue(isExistSecondLpCart, "Le lpCart n'est pas crée");
        }

        public void Test_Initialize_Create_LoadingPlan_with_Lp_CardBob()
        {
            // Prepare 2 Lpcart
            string name = TestContext.Properties["LpCartNamebob"].ToString();
            string code = TestContext.Properties["LpCartCodebob"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();
            string customerLp = "SMARTWINGS, A.S. (TVS)";
            string guestName = TestContext.Properties["GuestNameBob"].ToString();

            string nameBis = TestContext.Properties["LpCartNameBis"].ToString();
            string codeBis = TestContext.Properties["LpCartCodeBis"].ToString();

            //Prepare Loading plan
            string loadingPlanName = TestContext.Properties["LoadingPlanBob"].ToString();
            string loadingPlanNameBis = TestContext.Properties["LoadingPlanBobBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string aircraft = "BCS1";
            string type = TestContext.Properties["TypeBob"].ToString();
            string serviceName = TestContext.Properties["ServiceBob"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act Loading
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create first Loading Plan
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //add leg and service bob 
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);

                var generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.SelectLpCart(code + " - " + name);

                loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.ClickLoadingPlanLPCartEditLabels(trolleyScheme);
                generalInformationsPage.BackToList();
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            }
            bool isExistloadingPlanName = loadingPlanPage.CheckTotalNumber() > 0;
            Assert.IsTrue(isExistloadingPlanName, $"Le Loading Plan:{loadingPlanName} n'est pas créé");
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);

            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanNameBis, customerLp, route, aircraft, site, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //Edit the first loading plan
                var generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.SelectLpCart(codeBis + " - " + nameBis);
                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);
            }
            bool isExistloadingPlanNameBis = loadingPlanPage.CheckTotalNumber() > 0;
            Assert.IsTrue(isExistloadingPlanNameBis, $"Le Loading Plan:{loadingPlanNameBis} n'est pas créé");
        }

        public void Test_Initialize_Change_AircraftBOBWithLpCartInLoadingPlanMany()
        {
            string lpcartForFlight = TestContext.Properties["LpCartNamebob"].ToString();
            string codeBoblp = TestContext.Properties["LpCartCodebob"].ToString();
            string lpcartbob = TestContext.Properties["LpCartNameBis"].ToString();
            string codeBob = TestContext.Properties["LpCartCodeBis"].ToString();
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Bob comment";
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();

            string customerLp = "SMARTWINGS, A.S. (TVS)";
            string guestName = TestContext.Properties["GuestNameBob"].ToString();

            //Prepare Loading plan
            string loadingPlanName = TestContext.Properties["LoadingPlanBob"].ToString();
            string loadingPlanNameBis = TestContext.Properties["LoadingPlanBobBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string type = TestContext.Properties["TypeBob"].ToString();
            string serviceName = TestContext.Properties["ServiceBob"].ToString();

            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string aircraftBob = "BCS1";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Prepare
            //Créer 2 LPcart
            var lpCartPage = homePage.GoToFlights_LpCartPage();

            lpCartPage.Filter(LpCartPage.FilterType.Search, lpcartForFlight);

            if (lpCartPage.CheckTotalNumber() == 0)
            {
                // Create
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(codeBoblp, lpcartForFlight, site, customer, aircraftBob, from, to, comment);

                //Acces Onglet Cart                
                LpCartDetailPage.ClickAddtrolley();
                LpCartDetailPage.AddTrolley(trolleyName);

                //Create LpCart Scheme
                var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");
                LpCartDetailPage.BackToList();
            }
            else
            {
                var lpCartDetailPage = lpCartPage.ClickFirstLpCart();
                // On repousse la date du Lpcart
                var infomationPage = lpCartDetailPage.LpCartGeneralInformationPage();
                infomationPage.AddAircraf(aircraftBob);
                infomationPage.UpdateToDateTo();
                lpCartDetailPage.BackToList();
            }
            // Assert
            Assert.IsTrue(lpCartPage.CheckTotalNumber() >= 1, "Le lpCart n'est pas crée");

            lpCartPage.Filter(LpCartPage.FilterType.Search, lpcartbob);

            if (lpCartPage.CheckTotalNumber() == 0)
            {
                //Create second LpCart
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(codeBob, lpcartbob, site, customer, aircraftBob, from, to, comment);
                LpCartDetailPage.BackToList();
            }
            else
            {
                var lpCartDetailPage = lpCartPage.ClickFirstLpCart();
                // On repousse la date du Lpcart
                var infomationPage = lpCartDetailPage.LpCartGeneralInformationPage();
                infomationPage.AddAircraf(aircraftBob);
                infomationPage.UpdateToDateTo();
                lpCartDetailPage.BackToList();
            }
            // Assert
            Assert.IsTrue(lpCartPage.CheckTotalNumber() > 0, "Le lpCart n'est pas crée");

            //Créer 2 loading Plan
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create first Loading Plan
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraftBob, site, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //add leg and service bob 
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);

                var generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.SelectLpCart(codeBoblp + " - " + lpcartForFlight);

                loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.ClickLoadingPlanLPCartEditLabels(trolleyScheme);
                generalInformationsPage.BackToList();
            }

            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);

            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanNameBis, customerLp, route, aircraftBob, site, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //Edit the first loading plan
                var generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.SelectLpCart(codeBob + " - " + lpcartbob);
                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            }

            // Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            var editModalFlight = editPage.EditFlightInModal();
            editModalFlight.ChangeAircaftForLoadingPlan(aircraftBob);
            editModalFlight.SelectAllLoadingPlan();

            var isDecteted = editModalFlight.IsLpCartsDetectedInSelectedLoadingPlan();

            //Assert
            Assert.IsTrue(isDecteted, "Le message d'information n'apparait pas");
            //attention penser à desactiver le loading Plan en lançant le test FL_FLIG_Change_AircraftBOBWithLpCartInLoadingPlan juste après
        }

        public void Test_Initialize_Creat_flight(string aircraft = null)
        {
            string siteFrom =  TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string aircraft_from_sitting = TestContext.Properties["Aircraft"].ToString();
            // Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            var flightCreateModalpage1 = flightPage.FlightCreatePage();
            if ( aircraft != null)
            {
                flightCreateModalpage1.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo, null, "01", "22");
            }
            else
            {
                flightCreateModalpage1.FillField_CreatNewFlight(flightNumber, customer, aircraft_from_sitting, siteFrom, siteTo, null, "01", "22");
            }
            
            // Assert pour la verification de flight 

        }
        public void Test_Initialize__Add_Multiple_Flights()
        {
            // Prepare
            //string siteFrom = TestContext.Properties["Site"].ToString();
            string siteFrom = "ACE";
            string siteTo = "BCN";
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            //string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightMultiplModalpage = flightPage.CreateMultiFlights();
            flightMultiplModalpage.FillField_CreatMultiFlight(siteFrom, flightNumber, siteTo, aircraft);
            flightPage = flightMultiplModalpage.Validate();
        }
        public void Test_Initialize_Creat_Cancelled_flight(string aircraft = null)
        {
            string cancelled = "Cancelled";
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string aircraft_from_sitting = TestContext.Properties["Aircraft"].ToString();
            // Arrange
            var homePage = LogInAsAdmin();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            var flightCreateModalpage1 = flightPage.FlightCreatePage();
            if (aircraft != null)
            {
                flightCreateModalpage1.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo, null, "01", "22");
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.SetNewState("C");
                flightPage.ShowStatusMenu();
                flightPage.Filter(FlightPage.FilterType.Status, cancelled);
            }
            else
            {
                flightCreateModalpage1.FillField_CreatNewFlight(flightNumber, customer, aircraft_from_sitting, siteFrom, siteTo, null, "01", "22");
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.SetNewState("C");
                flightPage.ShowStatusMenu();
            }

            // Assert pour la verification de flight 

        }
        public string Test_Initialize_Creat_AirCraft()
        {
            string aircraftName = "Aircarft" + new Random().Next();
            string equipementType = "UNKNOWN";
        
            //Arrange
            var homePage = LogInAsAdmin();
            var parametersFlightPage = homePage.GoToParametres_Flights();
            var aircraftPage = parametersFlightPage.GoToAircraft();
            
            // cree un aircraft avec euipement ATLAS
            aircraftPage.CreateAircraftModelPage();
            aircraftPage.FillFieldsAircraftModelPage(aircraftName, equipementType);
            aircraftPage.Save();
            // Assert pour la creation 
            return aircraftName; 


            }
        public void Test_Initialize_Customer_Price_service()
        {
            // Prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["Customer"].ToString();
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);
            
            // Arrange
            HomePage homePage = LogInAsAdmin();
            
            //Creation service
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
            servicePricePage.BackToList();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
           }
        #endregion

        [Priority(0)]
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_VerifyServicesForFlight()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string serviceCode = "12345";
            string flightType = "International";

            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);

            string customer = customerLp.Substring(0, customerLp.IndexOf("("));

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var tabletPage = homePage.GoToParametres_Tablet();
            tabletPage.ClickFlightTypeTab();
            //check if flight type international exist
            var exist = tabletPage.VerifyInternationalExistOnSite(flightType, site);
            // if no create the flight type international
            if (!exist)
            {
                tabletPage.CreateNewFlightType(flightType, "Orange", site);
            }
            // else set the color gray
            tabletPage.EditFlightTypeColor(flightType, "Orange", site);

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(site, customer, fromDate, toDate);
                servicePage = pricePage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceName), $"Le service {serviceName} n'existe pas.");
        }

        //Update la date d'expiration d'un service BOB  
        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_VerifyServicesForFlightBOB()
        {
            // Prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["ServiceBob"].ToString();

            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(30);


            string customer = customerLp.Substring(0, customerLp.IndexOf(" ("));

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, "BUY ON BOARD");
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(customer, site, fromDate, toDate);
                servicePage = pricePage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceName), $"Le service {serviceName} n'existe pas.");
        }

        [TestMethod]
        [Priority(2)]
        [Timeout(_timeout)]
        public void FL_FLIG_Create_Lp_CardBob()
        {
            // Prepare
            string name = TestContext.Properties["LpCartNamebob"].ToString();
            string code = TestContext.Properties["LpCartCodebob"].ToString();

            string nameBis = TestContext.Properties["LpCartNameBis"].ToString();
            string codeBis = TestContext.Properties["LpCartCodeBis"].ToString();

            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string aircraft = "AB310";
            DateTime from = DateUtils.Now.AddDays(-1);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Bob comment";

            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();

            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            lpCartPage.Filter(LpCartPage.FilterType.StartDate, DateTime.Today.AddMonths(-1).AddDays(-11));
            lpCartPage.Filter(LpCartPage.FilterType.EndDate, DateTime.Today.AddMonths(2));

            if (lpCartPage.CheckTotalNumber() == 0)
            {
                // Create
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);

                //Acces Onglet Cart                
                LpCartDetailPage.ClickAddtrolley();
                LpCartDetailPage.AddTrolley(trolleyName);

                //Create LpCart Scheme
                var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");
                LpCartDetailPage.BackToList();
            }
            else
            {
                var lpCartDetailPage = lpCartPage.ClickFirstLpCart();
                // On repousse la date du Lpcart
                var infomationPage = lpCartDetailPage.LpCartGeneralInformationPage();
                infomationPage.UpdateToDateTo();
                lpCartDetailPage.BackToList();
            }
            // Assert
            bool isExistFirstNewLpCart = lpCartPage.CheckTotalNumber() >= 1;
            Assert.IsTrue(isExistFirstNewLpCart, "Le lpCart n'est pas crée");

            lpCartPage.Filter(LpCartPage.FilterType.Search, nameBis);

            if (lpCartPage.CheckTotalNumber() == 0)
            {
                //Create second LpCart
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(codeBis, nameBis, site, customer, aircraft, from, to, comment);
                LpCartDetailPage.BackToList();
            }
            else
            {
                var lpCartDetailPage = lpCartPage.ClickFirstLpCart();
                // On repousse la date du Lpcart
                var infomationPage = lpCartDetailPage.LpCartGeneralInformationPage();
                infomationPage.UpdateToDateTo();
                lpCartDetailPage.BackToList();
            }
            // Assert
            bool isExistSecondLpCart = lpCartPage.CheckTotalNumber() > 0;
            Assert.IsTrue(isExistSecondLpCart, "Le lpCart n'est pas crée");
        }

        //A surveiller 
        [TestMethod]
        [Priority(3)]
        [Timeout(_timeout)]
        public void FL_FLIG_Create_LoadingPlan_with_Lp_CardBob()
        {
            // Prepare 2 Lpcart
            string name = TestContext.Properties["LpCartNamebob"].ToString();
            string code = TestContext.Properties["LpCartCodebob"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();
            string customerLp = "SMARTWINGS, A.S. (TVS)";
            string guestName = TestContext.Properties["GuestNameBob"].ToString();

            string nameBis = TestContext.Properties["LpCartNameBis"].ToString();
            string codeBis = TestContext.Properties["LpCartCodeBis"].ToString();

            //Prepare Loading plan
            string loadingPlanName = TestContext.Properties["LoadingPlanBob"].ToString();
            string loadingPlanNameBis = TestContext.Properties["LoadingPlanBobBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string aircraft = "BCS1";
            string type = TestContext.Properties["TypeBob"].ToString();
            string serviceName = TestContext.Properties["ServiceBob"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act Loading
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create first Loading Plan
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //add leg and service bob 
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);

                var generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.SelectLpCart(code + " - " + name);

                loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.ClickLoadingPlanLPCartEditLabels(trolleyScheme);
                generalInformationsPage.BackToList();
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            }
            bool isExistloadingPlanName = loadingPlanPage.CheckTotalNumber() > 0;
            Assert.IsTrue(isExistloadingPlanName, $"Le Loading Plan:{loadingPlanName} n'est pas créé");
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);

            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanNameBis, customerLp, route, aircraft, site, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //Edit the first loading plan
                var generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.SelectLpCart(codeBis + " - " + nameBis);
                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);
            }
            bool isExistloadingPlanNameBis = loadingPlanPage.CheckTotalNumber() > 0;
            Assert.IsTrue(isExistloadingPlanNameBis, $"Le Loading Plan:{loadingPlanNameBis} n'est pas créé");
        }

        [TestMethod]
        [Priority(4)]
        [Timeout(_timeout)]
        public void FL_FLIG_Change_AircraftBOBWithLpCartInLoadingPlanMany()
        {
            string lpcartForFlight = TestContext.Properties["LpCartNamebob"].ToString();
            string codeBoblp = TestContext.Properties["LpCartCodebob"].ToString();
            string lpcartbob = TestContext.Properties["LpCartNameBis"].ToString();
            string codeBob = TestContext.Properties["LpCartCodeBis"].ToString();
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Bob comment";
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();

            string customerLp = "SMARTWINGS, A.S. (TVS)";
            string guestName = TestContext.Properties["GuestNameBob"].ToString();

            //Prepare Loading plan
            string loadingPlanName = TestContext.Properties["LoadingPlanBob"].ToString();
            string loadingPlanNameBis = TestContext.Properties["LoadingPlanBobBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string type = TestContext.Properties["TypeBob"].ToString();
            string serviceName = TestContext.Properties["ServiceBob"].ToString();

            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string aircraftBob = "BCS1";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Prepare
            //Créer 2 LPcart
            var lpCartPage = homePage.GoToFlights_LpCartPage();

            lpCartPage.Filter(LpCartPage.FilterType.Search, lpcartForFlight);

            if (lpCartPage.CheckTotalNumber() == 0)
            {
                // Create
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(codeBoblp, lpcartForFlight, site, customer, aircraftBob, from, to, comment);

                //Acces Onglet Cart                
                LpCartDetailPage.ClickAddtrolley();
                LpCartDetailPage.AddTrolley(trolleyName);

                //Create LpCart Scheme
                var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");
                LpCartDetailPage.BackToList();
            }
            else
            {
                var lpCartDetailPage = lpCartPage.ClickFirstLpCart();
                // On repousse la date du Lpcart
                var infomationPage = lpCartDetailPage.LpCartGeneralInformationPage();
                infomationPage.AddAircraf(aircraftBob);
                infomationPage.UpdateToDateTo();
                lpCartDetailPage.BackToList();
            }
            // Assert
            Assert.IsTrue(lpCartPage.CheckTotalNumber() >= 1, "Le lpCart n'est pas crée");

            lpCartPage.Filter(LpCartPage.FilterType.Search, lpcartbob);

            if (lpCartPage.CheckTotalNumber() == 0)
            {
                //Create second LpCart
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(codeBob, lpcartbob, site, customer, aircraftBob, from, to, comment);
                LpCartDetailPage.BackToList();
            }
            else
            {
                var lpCartDetailPage = lpCartPage.ClickFirstLpCart();
                // On repousse la date du Lpcart
                var infomationPage = lpCartDetailPage.LpCartGeneralInformationPage();
                infomationPage.AddAircraf(aircraftBob);
                infomationPage.UpdateToDateTo();
                lpCartDetailPage.BackToList();
            }
            // Assert
            Assert.IsTrue(lpCartPage.CheckTotalNumber() > 0, "Le lpCart n'est pas crée");

            //Créer 2 loading Plan
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create first Loading Plan
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraftBob, site, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //add leg and service bob 
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);

                var generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.SelectLpCart(codeBoblp + " - " + lpcartForFlight);

                loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.ClickLoadingPlanLPCartEditLabels(trolleyScheme);
                generalInformationsPage.BackToList();
            }

            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);

            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanNameBis, customerLp, route, aircraftBob, site, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //Edit the first loading plan
                var generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.SelectLpCart(codeBob + " - " + lpcartbob);
                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            }

            // Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            var editModalFlight = editPage.EditFlightInModal();
            editModalFlight.ChangeAircaftForLoadingPlan(aircraftBob);
            editModalFlight.SelectAllLoadingPlan();

            var isDecteted = editModalFlight.IsLpCartsDetectedInSelectedLoadingPlan();

            //Assert
            Assert.IsTrue(isDecteted, "Le message d'information n'apparait pas");
            //attention penser à desactiver le loading Plan en lançant le test FL_FLIG_Change_AircraftBOBWithLpCartInLoadingPlan juste après
        }

        

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Create_New_LoadingPlanBobForFlight()
        {
            // Prepare
            string aircraft = "B717";
            string aircraftBis = "B777";
            string site = TestContext.Properties["SiteLP"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string loadingPlanName = "Lp for flight bob";
            string loadingPlanNameBis = "Lp for flight bob bis";
            string type = TestContext.Properties["TypeBob"].ToString();
            string serviceName = TestContext.Properties["ServiceBob"].ToString();
            string guestName = "BOB";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.StartDate, DateUtils.Now.AddDays(-30));


            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                // Create
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBOBBtn();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);

                loadingPlanPage = loadingPlanDetailsPage.BackToList();
            }
            else
            {
                var editLoadindPlan = loadingPlanPage.ClickOnFirstLoadingPlan();
                editLoadindPlan.EditLoadingPlanInformations(DateUtils.Now.AddDays(30));
                editLoadindPlan.BackToList();
            }


            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            //Assert
            Assert.AreEqual(loadingPlanName, loadingPlanPage.GetFirstLoadingPlanName(), "Aucun loading plan bob créé.");

            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.StartDate, DateUtils.Now.AddDays(-30));


            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                // Create
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanNameBis, customerLp, route, aircraftBis, site, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBOBBtn();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);

                loadingPlanPage = loadingPlanDetailsPage.BackToList();
            }
            else
            {
                var editLoadindPlan = loadingPlanPage.ClickOnFirstLoadingPlan();
                editLoadindPlan.EditLoadingPlanInformations(DateUtils.Now.AddDays(30));
                editLoadindPlan.BackToList();
            }


            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);
            //Assert
            Assert.AreEqual(loadingPlanNameBis, loadingPlanPage.GetFirstLoadingPlanName(), "Aucun loading plan bob créé.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Vision_CustomersBob()
        {
            //Prepare
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string noCustomerBob = TestContext.Properties["CustomerLPFlight"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Verifie si le customer est Buy on board
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customerBob);
            var customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            var icao = customerGeneralInformationsPage.GetIcao();

            if (!customerGeneralInformationsPage.IsBoBVisible())
            {
                customerGeneralInformationsPage.ActiveBuyOnBoard();
            }

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, noCustomerBob, aircraft, siteFrom, siteTo);


            flightPage.ResetFilter();
            flightPage.ClickCustomerBob();
            flightPage.Filter(FlightPage.FilterType.CustomerBob, customer);

            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(flightPage.VerifyCustomer(icao), MessageErreur.FILTRE_ERRONE, "CustomersBob");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Export_LoadingPlanBOB_New_Version()
        {
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string serviceName = TestContext.Properties["ServiceBob"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var flightNumber = new Random().Next().ToString();
            bool newVersion = true;

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ClearDownloads();
            flightPage.ResetFilter();

            flightPage.ClickCustomerBob();
            flightPage.Filter(FlightPage.FilterType.CustomerBob, customer);
            flightPage.Filter(FlightPage.FilterType.HideCancelledFlight, true);

            if (flightPage.CheckTotalNumber() == 0)
            {
                flightPage.ResetFilter();
                // partir du mode BOB
                flightPage.ResetFilter();

                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);

                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var editPage = flightPage.EditFirstFlight(flightNumber);
                editPage.AddGuestType();
                editPage.AddService(serviceName);
                flightPage = editPage.CloseViewDetails();
                flightPage.ClickCustomerBob();
                flightPage.Filter(FlightPage.FilterType.CustomerBob, customer);
                flightPage.Filter(FlightPage.FilterType.HideCancelledFlight, true);
            }

            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }

            int flightNumbersNotCancelWithLegs = flightPage.GetTotalNumberLegs() - flightPage.GetTotalSplittedFlight();

            flightPage.ExportFlights(ExportType.ExportLoadingBob, newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Loading BOB", filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.AreEqual(resultNumber, flightNumbersNotCancelWithLegs, MessageErreur.EXCEL_DONNEES_KO);

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Vision_Customers_Print_DeliveryBOB_New_Version()
        {
            //Prepare
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string serviceName = TestContext.Properties["ServiceBob"].ToString();

            var flightNumber = new Random().Next().ToString();
            bool versionprint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Verifie si le customer est Buy on board
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customerBob);
            var customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            if (!customerGeneralInformationsPage.IsBoBVisible())
            {
                customerGeneralInformationsPage.ActiveBuyOnBoard();
            }

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            flightPage.ClearDownloads();

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType();
            editPage.AddService(serviceName);
            flightPage = editPage.CloseViewDetails();

            // Switch new flight state to V
            flightPage.SetNewState("V");
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(flightPage.GetFirstFlightStatus("V"), "Le statut du vol n'a pas été modifié à 'V' via la page index.");

            flightPage.ResetFilter();
            flightPage.ClickCustomerBob();
            flightPage.Filter(FlightPage.FilterType.CustomerBob, customer);

            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }

            // Print delivery note
            var reportPage = flightPage.PrintReport(PrintType.DeliveryBOB, versionprint);
            bool isPrintDeveliryNote = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isPrintDeveliryNote, "Le rapport n'a pas été généré.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Change_AircraftBOB()
        {
            //Prepere
            //prérequis : FL_FLIG_Create_New_LoadingPlanBobForFlight
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string aircraftForLp = "B717";
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString(); //TVS
            string flightNumber = new Random().Next().ToString();
            string lpCartNamebob = "Lp for flight bob";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.Filter(FlightPage.FilterType.Customer, customerBob);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            var popupTitle = editPage.ChangeAircaftForLoadingPlanAndReturnPoputTitle(aircraftForLp);
            Assert.AreEqual("INFORMATION (1)", popupTitle, "Pas de popup information");
            var popupSubmit = editPage.WaitForElementIsVisible(By.Id("dataAlertCancel"));
            popupSubmit.Click();
            Assert.IsTrue(editPage.IsLoadingPlanLoaded(lpCartNamebob), "Le loading plan n'est pas chargé.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Change_AircraftBOBWithLpCartInLoadingPlan()
        {
            //Prepare Flight
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string aircraftBob = "BCS1";
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string customerLp = "SMARTWINGS, A.S. (TVS)";
            string guestName = TestContext.Properties["GuestNameBob"].ToString();


            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //On verifie le guestType for return dans le customer
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customerLp);
            var customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            var customerBoBPage = customerGeneralInformationsPage.ClickBobTab();
            customerBoBPage.SelectGuestValueOnBob(guestName);


            // Act Flight
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.ChangeAircaftForLoadingPlan(aircraftBob);

            //assert
            Assert.IsTrue(editPage.IsLoadingPlanLoadedWithTrolley(), "Le loading plan ne sait pas automatiquement chargé selon aircraft.");
            //Désactive LpCart dans le flight  
            homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.EditFirstFlight(flightNumber);

            editPage.SelectLpCart("None");
            editPage.ChangeAircaftForLoadingPlan("-");
        }



        [TestMethod]

        [Timeout(_timeout)]
        public void FL_FLIG_Save_Flight_Only()
        {
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();


            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act Flight
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, siteFrom);
            flightPage.SetDateState(DateUtils.Now);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, "-", siteFrom, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            var beforeResult = editPage.CountServicesFromBob();
            var editModalFlight = editPage.EditFlightInModal();
            editModalFlight.ChangeAircaftForLoadingPlan("B717");
            editModalFlight.SaveForFlightOnly();
            var afterResult = editPage.CountServicesFromBob();
            //Assert
            Assert.AreNotEqual(beforeResult, afterResult, MessageErreur.SERVICE_NON_AJOUTE);
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_RemoveAllExistingService()
        {
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string qtyGuest = "10";

            var homePage = LogInAsAdmin();
            // Act Flight
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var editPage = flightPage.EditFirstFlight(flightNumber);
                //Edit flight
                editPage.AddGuestType("BOB");
                editPage.AddService(null);
                editPage.SetFinalQty(qtyGuest);
                editPage.CloseViewDetails();
                flightPage.EditFirstFlight(flightNumber);
                var editModalFlight = editPage.EditFlightInModal();
                editModalFlight.UnCheckRemoveAll();
                //Assert
                var serviceFinalQty = editPage.GetServiceFinalQty1(qtyGuest);
                Assert.IsTrue(serviceFinalQty, "Les quantités sont impactés");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_RemoveAllExistingServiceWithItem()
        {
            // Prepare
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string serviceName2 = TestContext.Properties["ServiceNameLP"].ToString();
            string qtyGuest = "10";
            string rnd = new Random().Next().ToString();
            string service = "service" + rnd;
            DateTime serviceDateFrom = DateUtils.Now;
            DateTime serviceDateTo = DateUtils.Now.AddMonths(2);
            string loadingPlanName = "lp" + rnd;
            string type = "BuyOnBoard";
            string route = TestContext.Properties["RouteLP"].ToString();
            string etaHours = "00";
            string etdHours = "23";
            string guestName = TestContext.Properties["GuestNameBob"].ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");

            //Arrange
            var homePage = LogInAsAdmin();
            // Create Service
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            var ServiceCreatePage = servicePage.ServiceCreatePage();
            ServiceCreatePage.FillFields_CreateServiceModalPage(service);
            var serviceGeneralInformationsPage = ServiceCreatePage.Create();
            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(siteFrom, customerBob, serviceDateFrom, serviceDateTo);
            DateTime lpDateFrom = serviceDateFrom.AddDays(2);
            DateTime lpDateTo = serviceDateTo.AddDays(-2);
            // Create LoadingPlan
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
            loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerBob, route, aircraft, siteFrom, type);
            loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(lpDateTo, lpDateFrom);
            var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
            //add leg and service bob 
            loadingPlanDetailsPage.ClickAddGuestBtn();
            loadingPlanDetailsPage.SelectGuest(guestName);
            loadingPlanDetailsPage.ClickCreateGuestBtn();
            //Add service
            loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
            loadingPlanDetailsPage.AddServiceBtn();
            loadingPlanDetailsPage.AddNewService(service);
            // Act Flight
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo, loadingPlanName, etaHours, etdHours, null, lpDateFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.SetDateState(lpDateFrom);
                var editPage = flightPage.EditFirstFlight(flightNumber);
                //Edit the first flight
                editPage.SelectGuestType(guestName);
                editPage.AddService(serviceName2);
                editPage.WaitPageLoading();
                editPage.SetFinalQty(qtyGuest);
                editPage.CloseViewDetails();
                flightPage.EditFirstFlight(flightNumber);
                editPage.WaitPageLoading();
                var editModalFlight = editPage.EditFlightInModal();
                editModalFlight.UnCheckRemoveAll();
                //Assert
                var serviceFinalQty = editPage.GetServiceFinalQty1(qtyGuest);
                Assert.IsTrue(serviceFinalQty, "Les quantités sont impactés");
                var countService = editPage.CountServicesFromBob();
                Assert.AreEqual(2, countService, MessageErreur.SERVICE_NON_AJOUTE);
            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
                //Delete LoadingPlan
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, null, null, lpDateTo.AddMonths(2));
                //Delete service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, service);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, deleteFrom);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, deleteTo);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_Search()
        {
            //Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customerICAO = TestContext.Properties["CustomerScheduleCode"].ToString();
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            Assert.AreEqual(flightNumber.ToString(), flightPage.GetFirstFlightNumber(), MessageErreur.FILTRE_ERRONE, "Search");

            flightPage.Filter(FlightPage.FilterType.SearchFlight, aircraft);

            var verifyAircraft = flightPage.VerifyAircraft(aircraft);
            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }
            Assert.IsTrue(verifyAircraft, MessageErreur.FILTRE_ERRONE);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, customerICAO);
            var verifyCustomerICAO = flightPage.VerifyCustomerICAO(customerICAO);
            Assert.IsTrue(verifyCustomerICAO, MessageErreur.FILTRE_ERRONE);


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_SortByAircraft()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.Filter(FlightPage.FilterType.SortBy, "aircraft");
            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }
            var isSortedByAircraft = flightPage.IsSortedByAircraft();
            Assert.IsTrue(isSortedByAircraft, MessageErreur.FILTRE_ERRONE, "Sort by aircraft");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_FilterTabletFlightType()
        {  // Prepare
            string international = "International";

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            var flightnumberbefore = flightPage.CheckTotalNumber();
            flightPage.Filter(FlightPage.FilterType.TabletFlight, international);
            var flightnumberAfter = flightPage.CheckTotalNumber();

            Assert.AreNotEqual(flightnumberbefore, flightnumberAfter, MessageErreur.FILTRE_ERRONE, "TABLET FLIGHT TYPE");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_UpdateDockNumber_WithImport()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersion = true;
            Random r = new Random();
            var docknumber1 = r.Next(1, 300).ToString();
            var docknumber2 = r.Next(1, 300).ToString();
            var docknumber3 = r.Next(1, 300).ToString();
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //login en tant que admin
            LogInAsAdmin();

            //naviguer sur homepage
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            //filtrer sur site ACE
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            //clean directory
            foreach (var file in taskDirectory.GetFiles())
            {
                file.Delete();
            }

            //Exporter le fichier
            flightPage.ExportFlights(ExportType.ExportDockNumberFile, newVersion);

            // On récupère le fichier du répertoire de téléchargement
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskDirectory.GetFiles());
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);

            // On récupère les 3 ids dans le fichier Excel
            var ids = OpenXmlExcel.GetValuesInList("Flight Id", "Flights", correctDownloadedFile.FullName);
            var id1 = ids[0];
            var id2 = ids[1];
            var id3 = ids[2];

            //On modifie le docknumber pour les 3 flights 
            OpenXmlExcel.WriteDataInCell("Dock number", "Flight Id", id1, "Flights", filePath, docknumber1, CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Dock number", "Flight Id", id2, "Flights", filePath, docknumber2, CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Dock number", "Flight Id", id3, "Flights", filePath, docknumber3, CellValues.SharedString);

            //On click sur le bouton d'import du dock number file
            flightPage.ImportFlights(ImportType.ImportDockNumbersFile, filePath);

            //On importe le fichier excel
            var serviceImport = new ServiceImportPage(WebDriver, TestContext);
            serviceImport.ImportFlightsExcelFile(filePath);

            //changer la page size a 100 item
            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }
            var result = flightPage.VerifyAllDockNumbersExist(docknumber1, docknumber2, docknumber3);
            Assert.IsTrue(result, "erreur lors de la mise a jour de la colonne dock number");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_UpdateDriver_WithImport()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string customer = "$$ - CAT Genérico";
            var newVersion = true;
            var driver1 = new Random().Next(1, 999).ToString();

            //login en tant que admin
            LogInAsAdmin();

            //naviguer sur homepage
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            //filtrer sur site ACE
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.ShowCustomersMenu();
            flightPage.Filter(FlightPage.FilterType.Customer, customer);

            if (!flightPage.isPageSizeEqualsTo100())
            {
                flightPage.PageSize("100");
            }

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            //clean directory
            flightPage.ClearDownloads();
            flightPage.PurgeDownloads();

            //Exporter le fichier

            flightPage.ExportFlights(ExportType.ExportDockNumberFile, newVersion);
            flightPage.ClosePrintButton();
            // On récupère le fichier du répertoire de téléchargement
            taskDirectory.Refresh();
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskDirectory.GetFiles());
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);

            // On récupère les 3 ids dans le fichier Excel
            var ids = OpenXmlExcel.GetValuesInList("Flight Number", "Flights", correctDownloadedFile.FullName);
            var id1 = ids[0];


            //On modifie le team pour les 3 flights 
            OpenXmlExcel.WriteDataInCell("Teams", "Flight Number", id1, "Flights", filePath, driver1, CellValues.SharedString);

            //On click sur le bouton d'import du dock number file
            flightPage.ImportFlights(ImportType.ImportDockNumbersFile, filePath);

            //On importe le fichier excel
            var serviceImport = new ServiceImportPage(WebDriver, TestContext);
            serviceImport.ImportFlightsExcelFile(filePath);

            //changer la page size a 100 item
            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("100");
            }

            // on recuperer la liste des docknumbers 
            flightPage.Filter(FilterType.SearchFlight, id1);
            var result = flightPage.VerifyAllDriversExist(driver1);
            Assert.IsTrue(result, "erreur lors de la mise a jour de la colonne driver");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_SortByDriver()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.Filter(FlightPage.FilterType.SortBy, "driver");
            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }
            var isSortedByDriver = flightPage.IsSortedByDriver();
            Assert.IsTrue(isSortedByDriver, MessageErreur.FILTRE_ERRONE, "Sort by driver");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_UpdateGate_WithImport()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            var newVersion = true;
            var gate1 = new Random().Next(1, 999).ToString();

            //login en tant que admin
            LogInAsAdmin();

            //naviguer sur homepage
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            //filtrer sur site ACE
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            if (!flightPage.isPageSizeEqualsTo100())
            {
                flightPage.PageSize("100");
            }

            // Gate
            flightPage.SetFirstGate("70");

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            //clean directory
            foreach (var file in taskDirectory.GetFiles())
            {
                file.Delete();
            }

            //Exporter le fichier
            flightPage.ExportFlights(ExportType.ExportDockNumberFile, newVersion);
            taskDirectory.Refresh();
            // On récupère le fichier du répertoire de téléchargement
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskDirectory.GetFiles());
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);

            // On récupère les 3 ids dans le fichier Excel
            var gates = OpenXmlExcel.GetValuesInList("Gate", "Flights", correctDownloadedFile.FullName);
            int i = 0;
            bool trouveGate = false;
            for (i = 0; i < gates.Count; i++)
            {
                if (gates[i] != "")
                {
                    trouveGate = true;
                    break;
                }
            }
            Assert.IsTrue(trouveGate, "pas de Gate dans les données du fichier Excel");
            var ids = OpenXmlExcel.GetValuesInList("Flight Number", "Flights", correctDownloadedFile.FullName);
            var id1 = ids[i];

            //On modifie le team pour les 3 flights 
            OpenXmlExcel.WriteDataInCell("Gate", "Flight Number", id1, "Flights", filePath, gate1, CellValues.SharedString);

            //On click sur le bouton d'import du dock number file
            flightPage.ImportFlights(ImportType.ImportDockNumbersFile, filePath);

            //On importe le fichier excel
            var serviceImport = new ServiceImportPage(WebDriver, TestContext);
            serviceImport.ImportFlightsExcelFile(filePath);

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            //changer la page size a 100 item
            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("100");
            }

            // on recuperer la liste des gates 
            flightPage.Filter(FilterType.SearchFlight, id1);
            var result = flightPage.VerifyAllGatesExist(gate1);
            Assert.IsTrue(result, "erreur lors de la mise a jour de la colonne gate");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_SortByGate()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.Filter(FlightPage.FilterType.SortBy, "gate");
            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }
            var isSortedByGate = flightPage.IsSortedByGate();
            Assert.IsTrue(isSortedByGate, MessageErreur.FILTRE_ERRONE, "Sort by gate");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Registration_Manuel_Export_Import()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            var newVersion = true;
            var registration1 = "111";
            var registration2 = "222";
            var registration3 = "333";

            //login en tant que admin
            LogInAsAdmin();

            //naviguer sur homepage
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            //filtrer sur site ACE
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            //clean directory
            foreach (var file in taskDirectory.GetFiles())
            {
                file.Delete();
            }

            //Exporter le fichier
            flightPage.ExportFlights(ExportType.ExportDockNumberFile, newVersion);

            // On récupère le fichier du répertoire de téléchargement
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskDirectory.GetFiles());
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);

            // On récupère les 3 ids dans le fichier Excel
            var ids = OpenXmlExcel.GetValuesInList("Flight Id", "Flights", correctDownloadedFile.FullName);
            var id1 = ids[0];
            var id2 = ids[1];
            var id3 = ids[2];


            //On modifie le team pour les 3 flights 
            OpenXmlExcel.WriteDataInCell("Registration", "Flight Id", id1, "Flights", filePath, registration1, CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Registration", "Flight Id", id2, "Flights", filePath, registration2, CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Registration", "Flight Id", id3, "Flights", filePath, registration3, CellValues.SharedString);

            //On click sur le bouton d'import du dock number file
            flightPage.ImportFlights(ImportType.ImportDockNumbersFile, filePath);

            //On importe le fichier excel
            var serviceImport = new ServiceImportPage(WebDriver, TestContext);
            serviceImport.ImportFlightsExcelFile(filePath);

            //changer la page size a 100 item
            if (!flightPage.isPageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }

            // on recupere la liste des registrations : non mise à jour
            var result = flightPage.VerifyRegistrationsNotExist(registration1, registration2, registration3);
            Assert.IsTrue(result, "registration maj alors qu'il ne fallait pas");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_SortByRegistration()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.Filter(FlightPage.FilterType.SortBy, "registration");
            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }
            var isSortedByRegistration = flightPage.IsSortedByResgistration();
            Assert.IsTrue(isSortedByRegistration, MessageErreur.FILTRE_ERRONE, "Sort by registration");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_AddAndModifyDockNumber()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var newVersion = true;
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            var dockNumberToSet = new Random().Next(1, 300).ToString();


            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            // changer la page size a 100 item
            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }
            //click on the first line
            var line = WebDriver.FindElement(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]"));
            line.Click();
            //click sur first line
            var docknumber = WebDriver.FindElement(By.XPath("//*[@id=\"item_DockNumber\"]"));
            docknumber.Clear();
            docknumber.Click();
            docknumber.SendKeys(dockNumberToSet);
            //click on 2nd line
            WebDriver.FindElement(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]")).Click();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            flightPage.ClearDownloads();

            //Exporter le fichier
            flightPage.ExportFlights(ExportType.ExportDockNumberFile, newVersion);
            flightPage.WaiForLoad();
            // On récupère le fichier du répertoire de téléchargement
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskDirectory.GetFiles());
            Assert.IsNotNull(correctDownloadedFile);
            // on recupere les docknumber dans la liste
            var docknumbers = OpenXmlExcel.GetValuesInList("Dock number", "Flights", correctDownloadedFile.FullName);
            //verifie si le dock number existe
            var test = false;
            foreach (var number in docknumbers)
            {
                if (dockNumberToSet == number)
                {
                    test = true;
                }
            }
            //resultat
            if (!test)
            {
                Assert.Fail("pas de modification dans le fichier excel");
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_AddAndModifyDriver()
        {
            //Prepare
            var success = false;
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var newVersion = true;
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            var driverToSet = new Random().Next(1, 300).ToString();
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string olddriver = "70";

            //Arrange
            var homePage = LogInAsAdmin();

            //act
            FlightPage flightPage = homePage.GoToFlights_FlightPage();

            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                //Add driver
                flightPage.CliCkOnFirstFlight(2);
                flightPage.SetDriver(olddriver);
                string oldvalue = flightPage.GetIndexDriverValue();
                //Update driver 
                flightPage.SetDriver(driverToSet);
                string modifiedvalue = flightPage.GetIndexDriverValue();

                //Assert
                Assert.AreNotEqual(oldvalue, modifiedvalue, MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE);

                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                flightPage.ClearDownloads();
                //Exporter le fichier
                flightPage.ExportFlights(ExportType.ExportDockNumberFile, newVersion);
                flightPage.WaiForLoad();
                // On récupère le fichier du répertoire de téléchargement
                var correctDownloadedFile = flightPage.GetExportExcelFile(taskDirectory.GetFiles());
                Assert.IsNotNull(correctDownloadedFile);
                // on recupere les docknumber dans la liste
                var drivers = OpenXmlExcel.GetValuesInList("Teams", "Flights", correctDownloadedFile.FullName);
                //verifie si le dock number existe
                foreach (var dr in drivers)
                {
                    if (driverToSet == dr)
                    {
                        success = true;
                    }
                }
                //resultat
                if (!success)
                {
                    Assert.Fail("pas de modification dans le fichier excel");
                }
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.SetDateFromAndTo(DateTime.Now, DateTime.Now);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }

        //
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_AddAndModifyGate()
        {
            var success = false;
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var newVersion = true;
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            var gateToSet = new Random().Next(1, 300).ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }
            flightPage.CliCkOnFirstFlight(2);
            flightPage.SetFirstGate(gateToSet);
            flightPage.CliCkOnFirstFlight(3);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            flightPage.ClearDownloads();

            //Export excel file
            flightPage.ExportFlights(ExportType.ExportDockNumberFile, newVersion);
            flightPage.WaiForLoad();
            // On récupère le fichier du répertoire de téléchargement
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskDirectory.GetFiles());
            Assert.IsNotNull(correctDownloadedFile);
            // on recupere les gates
            var gates = OpenXmlExcel.GetValuesInList("Gate", "Flights", correctDownloadedFile.FullName);
            //verifie si gate exist
            foreach (var g in gates)
            {
                if (gateToSet == g)
                {
                    success = true;
                }
            }
            //resultat
            if (!success)
            {
                Assert.Fail("pas de modification dans le fichier excel");
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_AddAndModifyRegistration()
        {
            var success = false;
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var newVersion = true;
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            var registrationToSet = new Random().Next(1, 300).ToString();
            string status = "Preval";
            string columnName = "Registration";
            string sheetName = "Flights";

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.Filter(FlightPage.FilterType.Status, status);
            // changer la page size a 100 item
            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }
            flightPage.AddRegistrationOnFirstFlight(registrationToSet);
            flightPage.ClearDownloads();
            //Export excel file
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            flightPage.ExportFlights(ExportType.ExportDockNumberFile, newVersion);
            flightPage.WaiForLoad();
            // On récupère le fichier du répertoire de téléchargement
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskDirectory.GetFiles());
            Assert.IsNotNull(correctDownloadedFile);
            // on recupere les gates
            var registrations = OpenXmlExcel.GetValuesInList(columnName, sheetName, correctDownloadedFile.FullName);
            //verifie si gate exist
            foreach (var r in registrations)
            {
                if (registrationToSet == r)
                {
                    success = true;
                }
            }
            //resultat
            if (!success)
            {
                Assert.Fail("pas de modification dans le fichier excel");
            }
        }

        // 

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_Site()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            bool verifySite = flightPage.VerifySiteFrom(siteFrom);
            Assert.IsTrue(verifySite, MessageErreur.FILTRE_ERRONE, "Sites");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_SortByAirline()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string customer2 = TestContext.Properties["Bob_Customer"].ToString();
            string flightNumber1 = "Flight 1 - " + new Random().Next().ToString();
            string flightNumber2 = "Flight 2 - " + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var totalNumbe = flightPage.CheckTotalNumber();
            try
            {
                if (totalNumbe < 2)
                {
                    //create flight1
                    var flightCreateModalpage1 = flightPage.FlightCreatePage();
                    flightCreateModalpage1.FillField_CreatNewFlight(flightNumber1, customer, aircraft, siteFrom, siteTo, null, "01", "22");

                    if (totalNumbe < 1)
                    {
                        //create flight2
                        var flightCreateModalpage2 = flightPage.FlightCreatePage();
                        flightCreateModalpage2.FillField_CreatNewFlight(flightNumber2, customer2, aircraft, siteFrom, siteTo, null, "01", "22");
                    }
                }
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                if (!flightPage.PageSizeEqualsTo100())
                {
                    flightPage.PageSize("8");
                    flightPage.PageSize("100");
                }

                flightPage.Filter(FlightPage.FilterType.SortBy, "airline");
                var isSortedByAirline = flightPage.IsSortedByAirline();

                //Assert
                Assert.IsTrue(isSortedByAirline, MessageErreur.FILTRE_ERRONE, "Sort by airline");
            }
            finally
            {
                //delete Flight 
                if (totalNumbe < 2)
                {
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber1);
                    flightPage.MassiveDeleteMenus(flightNumber1, siteTo, customer, true);
                    flightPage.ResetFilter();

                    if (totalNumbe < 1)
                    {
                        flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber2);
                        flightPage.MassiveDeleteMenus(flightNumber2, siteTo, customer2, true);
                        flightPage.ResetFilter();
                    }
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_SortByFlightNo()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNumber1 = new Random().Next().ToString();
            string flightNumber2 = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var totalNumbe = flightPage.CheckTotalNumber();
            try
            {
                if (totalNumbe < 2)
                {
                    //create flight1
                    var flightCreateModalpage1 = flightPage.FlightCreatePage();
                    flightCreateModalpage1.FillField_CreatNewFlight(flightNumber1, customer, aircraft, siteFrom, siteTo, null, "01", "22");

                    if (totalNumbe < 1)
                    {
                        //create flight
                        var flightCreateModalpage2 = flightPage.FlightCreatePage();
                        flightCreateModalpage2.FillField_CreatNewFlight(flightNumber2, customer, aircraft, siteFrom, siteTo, null, "01", "22");
                    }
                }
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

                if (!flightPage.PageSizeEqualsTo100())
                {
                    flightPage.PageSize("8");
                    flightPage.PageSize("100");
                }

                flightPage.Filter(FlightPage.FilterType.SortBy, "flight no");
                var isSortedByFlightNo = flightPage.IsSortedByFlightNo();

                //Assert
                Assert.IsTrue(isSortedByFlightNo, MessageErreur.FILTRE_ERRONE, "Sort by flight no");

            }
            finally
            {
                //delete Flight 
                if (totalNumbe < 2)
                {
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber1);
                    flightPage.MassiveDeleteMenus(flightNumber1, siteTo, customer, true);
                    flightPage.ResetFilter();

                    if (totalNumbe < 1)
                    {
                        flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber2);
                        flightPage.MassiveDeleteMenus(flightNumber2, siteTo, customer, true);
                        flightPage.ResetFilter();
                    }
                }
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_SortByETD()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNumber1 = new Random().Next().ToString();
            string flightNumber2 = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var totalNumbe = flightPage.CheckTotalNumber();
            try
            {
                if (totalNumbe < 2)
                {
                    //create flight1
                    var flightCreateModalpage1 = flightPage.FlightCreatePage();
                    flightCreateModalpage1.FillField_CreatNewFlight(flightNumber1, customer, aircraft, siteFrom, siteTo, null, "01", "22");

                    if (totalNumbe < 1)
                    {
                        //create flight
                        var flightCreateModalpage2 = flightPage.FlightCreatePage();
                        flightCreateModalpage2.FillField_CreatNewFlight(flightNumber2, customer, aircraft, siteFrom, siteTo, null, "01", "22");
                    }
                }

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

                if (!flightPage.PageSizeEqualsTo100())
                {
                    flightPage.PageSize("8");
                    flightPage.PageSize("100");
                }

                flightPage.Filter(FlightPage.FilterType.SortBy, "ETD");
                var isSortedByEtd = flightPage.IsSortedByEtd();

                //Assert
                Assert.IsTrue(isSortedByEtd, MessageErreur.FILTRE_ERRONE, "Sort by ETD");
            }
            finally
            {
                //delete Flight 
                if (totalNumbe < 2)
                {
                    flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber1);
                    flightPage.MassiveDeleteMenus(flightNumber1, siteTo, customer, true);

                    if (totalNumbe < 1)
                    {
                        flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber2);
                        flightPage.MassiveDeleteMenus(flightNumber2, siteTo, customer, true);
                        flightPage.ResetFilter();
                    }
                }
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_SortByETA()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNumber1 = new Random().Next().ToString();
            string flightNumber2 = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var totalNumbe = flightPage.CheckTotalNumber();
            try
            {
                if (totalNumbe < 2)
                {
                    //create flight1
                    var flightCreateModalpage1 = flightPage.FlightCreatePage();
                    flightCreateModalpage1.FillField_CreatNewFlight(flightNumber1, customer, aircraft, siteFrom, siteTo, null, "01", "22");

                    if (totalNumbe < 1)
                    {
                        //create flight
                        var flightCreateModalpage2 = flightPage.FlightCreatePage();
                        flightCreateModalpage2.FillField_CreatNewFlight(flightNumber2, customer, aircraft, siteFrom, siteTo, null, "01", "22");
                    }
                }

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

                if (!flightPage.PageSizeEqualsTo100())
                {
                    flightPage.PageSize("8");
                    flightPage.PageSize("100");
                }

                flightPage.Filter(FlightPage.FilterType.SortBy, "ETA");
                var isSortedByEta = flightPage.IsSortedByEta();

                //Assert
                Assert.IsTrue(isSortedByEta, MessageErreur.FILTRE_ERRONE, "Sort by ETA");
            }
            finally
            {
                //delete Flight 
                if (totalNumbe < 2)
                {
                    flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber1);
                    flightPage.MassiveDeleteMenus(flightNumber1, siteTo, customer, true);

                    if (totalNumbe < 1)
                    {
                        flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber2);
                        flightPage.MassiveDeleteMenus(flightNumber2, siteTo, customer, true);
                        flightPage.ResetFilter();
                    }
                }
            }

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_SortByFrom()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteFrom2 = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNumber1 = "Flight 1 - " + new Random().Next().ToString();
            string flightNumber2 = "Flight 2 - " + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            var totalNumbe = flightPage.CheckTotalNumber();
            try
            {
                if (totalNumbe < 2)
                {
                    //create flight1
                    flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                    var flightCreateModalpage1 = flightPage.FlightCreatePage();
                    flightCreateModalpage1.FillField_CreatNewFlight(flightNumber1, customer, aircraft, siteFrom, siteTo, null, "01", "22");

                    if (totalNumbe < 1)
                    {
                        //create flight
                        var flightCreateModalpage2 = flightPage.FlightCreatePage();
                        flightCreateModalpage2.FillField_CreatNewFlight(flightNumber2, customer, aircraft, siteFrom, siteTo, null, "01", "22");
                    }
                }

                flightPage.ResetFilter();

                if (!flightPage.PageSizeEqualsTo100())
                {
                    flightPage.PageSize("8");
                    flightPage.PageSize("100");
                }
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SortBy, "from");
                var isSortedByFrom = flightPage.IsSortedByFrom();

                //Assert
                Assert.IsTrue(isSortedByFrom, MessageErreur.FILTRE_ERRONE, "Sort by from");
            }
            finally
            {
                //delete Flight 
                if (totalNumbe < 2)
                {
                    flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber1);
                    flightPage.MassiveDeleteMenus(flightNumber1, siteFrom, customer, true);
                    flightPage.ResetFilter();

                    if (totalNumbe < 1)
                    {
                        flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                        flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber2);
                        flightPage.MassiveDeleteMenus(flightNumber2, siteFrom, customer, true);
                        flightPage.ResetFilter();
                    }
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_Customers()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string customerCode = TestContext.Properties["CustomerScheduleCode"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            string customer = customerLp.Substring(0, customerLp.IndexOf("("));
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            flightPage.ShowCustomersMenu();
            flightPage.Filter(FlightPage.FilterType.Customer, customer);

            if (flightPage.CheckTotalNumber() < 20)
            {
                // Create
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

                flightPage.ShowCustomersMenu();
                flightPage.Filter(FlightPage.FilterType.Customer, customer);
            }

            if (!flightPage.PageSizeEqualsTo100())
            {
                flightPage.PageSize("8");
                flightPage.PageSize("100");
            }

            Assert.IsTrue(flightPage.VerifyCustomer(customerCode), MessageErreur.FILTRE_ERRONE, "Customers");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_Status()
        {

            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();

            string preval = "Preval";
            string valid = "Valid";
            string invoice = "Invoice";
            string cancelled = "Cancelled";
            string none = "None";

            string flightNumber = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();

            // Create
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber.ToString());
            var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.AddGuestType();
            editPage.AddService(serviceName);
            flightPage = editPage.CloseViewDetails();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber.ToString());

            flightPage.SetNewState("");
            WebDriver.Navigate().Refresh();
            flightPage.ShowStatusMenu();
            flightPage.Filter(FlightPage.FilterType.Status, none);
            string resultNone = flightPage.GetFirstFlightNumber();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber.ToString());

            flightPage.SetNewState("P");
            WebDriver.Navigate().Refresh();
            flightPage.ShowStatusMenu();
            flightPage.Filter(FlightPage.FilterType.Status, preval);
            string resultPreval = flightPage.GetFirstFlightNumber();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber.ToString());

            flightPage.SetNewState("V");
            WebDriver.Navigate().Refresh();
            flightPage.ShowStatusMenu();
            flightPage.Filter(FlightPage.FilterType.Status, valid);
            string resultValid = flightPage.GetFirstFlightNumber();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber.ToString());

            flightPage.SetNewState("I");
            WebDriver.Navigate().Refresh();
            flightPage.ShowStatusMenu();
            flightPage.Filter(FlightPage.FilterType.Status, invoice);
            string resultInvoice = flightPage.GetFirstFlightNumber();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber.ToString());

            flightPage.SetNewState("C");
            flightPage.ShowStatusMenu();
            flightPage.Filter(FlightPage.FilterType.Status, cancelled);
            string resultCancelled = flightPage.GetFirstFlightNumber();

            Assert.AreEqual(flightNumber.ToString(), resultNone, MessageErreur.FILTRE_ERRONE, none);
            Assert.AreEqual(flightNumber.ToString(), resultPreval, MessageErreur.FILTRE_ERRONE, preval);
            Assert.AreEqual(flightNumber.ToString(), resultValid, MessageErreur.FILTRE_ERRONE, valid);
            Assert.AreEqual(flightNumber.ToString(), resultInvoice, MessageErreur.FILTRE_ERRONE, invoice);
            Assert.AreEqual(flightNumber.ToString(), resultCancelled, MessageErreur.FILTRE_ERRONE, cancelled);
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_DayTime_ETDFrom()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string etaHours = "00";
            string etdHours = "05";

            //Arrange
            var homePage = LogInAsAdmin();

            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo, null, etaHours, etdHours);
            try
            {
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);

                flightPage.Filter(FlightPage.FilterType.ETDFrom, "0500AM");
                var displayedFlight1 = flightPage.GetFirstFlightNumber();
                Assert.AreEqual(flightNumber, displayedFlight1, MessageErreur.FILTRE_ERRONE, "ETD From");

                flightPage.Filter(FlightPage.FilterType.ETDFrom, "0400AM");
                var displayedFlight2 = flightPage.GetFirstFlightNumber();
                Assert.AreEqual(flightNumber, displayedFlight2, MessageErreur.FILTRE_ERRONE, "ETD From");

                flightPage.Filter(FlightPage.FilterType.ETDFrom, "0530AM");
                var isFlightDisplayed1 = flightPage.IsFlightDisplayed();
                Assert.IsFalse(isFlightDisplayed1, MessageErreur.FILTRE_ERRONE, "ETD From");

                flightPage.Filter(FlightPage.FilterType.ETDFrom, "0430AM");
                flightPage.Filter(FlightPage.FilterType.ETDTo, "0530AM");
                var displayedFlight3 = flightPage.GetFirstFlightNumber();
                Assert.AreEqual(flightNumber, displayedFlight3, MessageErreur.FILTRE_ERRONE, "ETD From et ETD To");

                flightPage.Filter(FlightPage.FilterType.ETDFrom, "0400AM");
                flightPage.Filter(FlightPage.FilterType.ETDTo, "0430AM");
                var isFlightDisplayed2 = flightPage.IsFlightDisplayed();
                Assert.IsFalse(isFlightDisplayed2, MessageErreur.FILTRE_ERRONE, "ETD From et ETD To");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_DayTime_ETDTo()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string etaHours = "13";
            string etdHours = "15";

            //Arrange
            var homePage = LogInAsAdmin();

            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo, null, etaHours, etdHours);
            try
            {
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);

                flightPage.Filter(FlightPage.FilterType.ETDTo, "0300PM");
                var displayedFlight1 = flightPage.GetFirstFlightNumber();
                Assert.AreEqual(flightNumber, displayedFlight1, MessageErreur.FILTRE_ERRONE, "ETD From");

                flightPage.Filter(FlightPage.FilterType.ETDTo, "0400PM");
                var displayedFlight2 = flightPage.GetFirstFlightNumber();
                Assert.AreEqual(flightNumber, displayedFlight2, MessageErreur.FILTRE_ERRONE, "ETD From");

                flightPage.Filter(FlightPage.FilterType.ETDTo, "0230PM");
                var isFlightDisplayed1 = flightPage.IsFlightDisplayed();
                Assert.IsFalse(isFlightDisplayed1, MessageErreur.FILTRE_ERRONE, "ETD From");

                flightPage.Filter(FlightPage.FilterType.ETDFrom, "0230PM");
                flightPage.Filter(FlightPage.FilterType.ETDTo, "0330PM");
                var displayedFlight3 = flightPage.GetFirstFlightNumber();
                Assert.AreEqual(flightNumber, displayedFlight3, MessageErreur.FILTRE_ERRONE, "ETD From et ETD To");

                flightPage.Filter(FlightPage.FilterType.ETDFrom, "0400PM");
                flightPage.Filter(FlightPage.FilterType.ETDTo, "0430AM");
                var isFlightDisplayed2 = flightPage.IsFlightDisplayed();
                Assert.IsFalse(isFlightDisplayed2, MessageErreur.FILTRE_ERRONE, "ETD From et ETD To");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_BatchTime_PrintExport()
        {
            //Attention, le vol ne sera pas sur l'export, s'il n'a pas de service actif dans un guest price
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Delivery Note_-_";
            string DocFileNameZipBegin = "All_files_";

            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string serviceName = TestContext.Properties["ServiceBob"].ToString();
            var flightNumber = new Random().Next().ToString();
            bool versionprint = true;

            string etdFrom = "0530AM";
            string etdTo = "0630AM";

            //Arrange
            var homePage = LogInAsAdmin();

            // Act
            var flightPage = homePage.GoToFlights_FlightPage();

            flightPage.ClearDownloads();
            try
            {
                //Create flight
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                Thread.Sleep(2000);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo, null, "00", "06");

                //Add Service
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var editPage = flightPage.EditFirstFlight(flightNumber);
                editPage.AddGuestType();
                editPage.AddService(serviceName);
                flightPage = editPage.CloseViewDetails();

                //Appliquer le filtre
                flightPage.ResetFilter();
                flightPage.ClickCustomerBob();
                flightPage.Filter(FlightPage.FilterType.CustomerBob, customer);
                flightPage.Filter(FlightPage.FilterType.HideCancelledFlight, true);
                flightPage.Filter(FlightPage.FilterType.ETDFrom, etdFrom);
                flightPage.Filter(FlightPage.FilterType.ETDTo, etdTo);

                var valuesVerified = flightPage.VerifyETDInRange(etdFrom, etdTo);
                //Assert
                Assert.IsTrue(valuesVerified, " Les résultats ne s'accordent pas au filtre appliqué.");

                ////Print delivery note
                //var reportPage = flightPage.PrintReport(PrintType.DeliveryNote, versionprint);
                //bool isPrintDeveliryNote = reportPage.IsReportGenerated();
                //reportPage.Close();

                ////Assert
                //Assert.IsTrue(isPrintDeveliryNote, "Le rapport n'a pas été généré.");
                //reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                //// cliquer sur All
                //string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
                //FileInfo filePdf = new FileInfo(trouve);
                //filePdf.Refresh();
                //Assert.IsTrue(filePdf.Exists, trouve + " non généré");

                ////Vérifier le bon nombre de page
                //Assert.AreEqual(flightPage.CheckTotalNumber(), flightPage.CheckPdfNumberOfPage(filePdf), "Il n'y a pas le bon nombre de vols affichés dans le document.");
                ////Vérifier la présence des mêmes id que sur winrest
                //flightPage.CheckPrint(filePdf);

                //// Print track sheet
                //flightPage.ClearDownloads();

                //flightPage.Filter(FlightPage.FilterType.HideCancelledFlight, true);
                //reportPage = flightPage.PrintReport(PrintType.TrackSheet, versionprint);
                //bool isPrintTrackSheet = reportPage.IsReportGenerated();
                //reportPage.Close();

                ////Assert
                //Assert.IsTrue(isPrintTrackSheet, "Le tracksheet n'a pas été généré.");
                //DocFileNamePdfBegin = "Print TrackSheet_-_";
                //reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                //// cliquer sur All
                //trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
                //filePdf = new FileInfo(trouve);
                //filePdf.Refresh();
                //Assert.IsTrue(filePdf.Exists, trouve + " non généré");
                //flightPage.CheckPrint(filePdf);
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
            
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_Hide_Cancelled_Flights()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            flightPage.SetNewState("C");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.Filter(FlightPage.FilterType.HideCancelledFlight, true);

            //Assert
            Assert.IsFalse(flightPage.IsFlightDisplayed(), "Le flight est affiché malgré le fait qu'il soit CANCELLED");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Create_New_Flight()
        {
            // Prepare
            string rnd = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString(); //B737
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string flightNumber = rnd;
            string service = "service" + rnd;
            string loadingPlanName = "lp" + rnd;
            string lpCartName = "lpcart" + rnd;
            string route = "ACE-BCN"; // TestContext.Properties["RouteLP"].ToString();
            string type = "BuyOnBoard";
            string code = "codeLp" + DateUtils.Now.ToString("dd/MM/yyyy") + rnd;
            string comment = "Bob comment";
            string etaHours = "00";
            string etdHours = "23";
            DateTime serviceDateFrom = DateUtils.Now;
            DateTime serviceDateTo = DateUtils.Now.AddMonths(2);
            string guestName = TestContext.Properties["GuestNameBob"].ToString();
            string serviceName = service;// TestContext.Properties["ServiceBob"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            // Create Service
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            var ServiceCreatePage = servicePage.ServiceCreatePage();
            ServiceCreatePage.FillFields_CreateServiceModalPage(service);
            var serviceGeneralInformationsPage = ServiceCreatePage.Create();
            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(siteFrom, customer, serviceDateFrom, serviceDateTo);
            DateTime lpDateFrom = serviceDateFrom.AddDays(2);
            DateTime lpDateTo = serviceDateTo.AddDays(-2);

            // Create LpCart
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            var lpCartModalCreate = lpCartPage.LpCartCreatePage();
            var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCartWithRoutes(code, lpCartName, siteFrom, customer, aircraft, lpDateFrom, lpDateTo, comment, route);

            // Create LoadingPlan
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
            loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customer, route, aircraft, siteFrom, type);
            loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(lpDateTo, lpDateFrom);
            var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
            //add leg and service bob 
            loadingPlanDetailsPage.ClickAddGuestBtn();
            loadingPlanDetailsPage.SelectGuest(guestName);
            loadingPlanDetailsPage.ClickCreateGuestBtn();
            loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
            loadingPlanDetailsPage.AddServiceBtn();
            loadingPlanDetailsPage.AddNewService(serviceName);


            // Create Flight
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, lpDateFrom);
            try
            {
                var displayedNumber = flightPage.GetFirstFlightNumber();
                Assert.AreEqual(flightNumber, displayedNumber, "Le vol n'a pas été créé.");
                var flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
                var lPCartIsExist = flightDetailsPage.LPCartIsExist(lpCartName);
                Assert.IsTrue(lPCartIsExist, "Le LPcart rattaché avec le LP n'est pas ajouté avec le vol");
            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();

                //Delete LoadingPlan
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, null, null, lpDateTo.AddMonths(2));

                //Delete LpCart
                lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                lpCartPage.DeleteLpCart();

                //Delete service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, service);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Create_New_Flight_Error()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(null, null, aircraft, siteFrom, siteTo);
            //Assert
            var isValid = flightCreateModalpage.FlightIsValid();
            Assert.IsFalse(isValid, "Le flight a été créé malgré l'erreur.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Add_Multiple_Flights()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightMultiplModalpage = flightPage.CreateMultiFlights();
            flightMultiplModalpage.FillField_CreatMultiFlight(siteFrom, flightNumber, siteTo, aircraft);

            flightPage = flightMultiplModalpage.Validate();

            // On vérifie que le vol est créé pour tous les jours de la semaine
            Assert.IsTrue(flightPage.IsMultiFlightAdded(flightNumber), "Les flights n'ont pas été ajoutés.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Add_Multiple_Flights_No_Rules()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightMultiplModalpage = flightPage.CreateMultiFlights();
            flightMultiplModalpage.Validate();
            var isErrorDisplayed = flightMultiplModalpage.IsErrorMessage();
            Assert.IsTrue(isErrorDisplayed, "Aucun message d'erreur n'est apparu lors de la création du multiFlight.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Add_Multiple_Flights_Duplicate()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightMultiplModalpage = flightPage.CreateMultiFlights();
            flightMultiplModalpage.FillField_CreatMultiFlight(siteFrom, flightNumber, siteTo, aircraft);
            flightMultiplModalpage.DuplicateFlight();
            var verifyNumber = flightMultiplModalpage.CompareDuplicatedFlightNumber();
            Assert.IsTrue(verifyNumber, "Le flight number n'a pas été dupliqué.");
            var verifyAircraft = flightMultiplModalpage.CompareAircraft();
            Assert.IsTrue(verifyAircraft, "L'aircraft n'a pas été dupliqué.");
            var verifyDays = flightMultiplModalpage.CompareDays();
            Assert.IsTrue(verifyDays, "Les jours de vol n'ont pas été dupliqués.");
            var verifyDate = flightMultiplModalpage.CompareDate();
            Assert.IsTrue(verifyDate, "Les dates du vol n'ont pas été dupliquées.");
            var verifyAirport = flightMultiplModalpage.CompareDuplicatedAirport();
            Assert.IsTrue(verifyAirport, "Les airports du vol n'ont pas été dupliqués.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Add_Multiple_Flights_RemoveFlight()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightMultiplModalpage = flightPage.CreateMultiFlights();
            flightMultiplModalpage.FillField_CreatMultiFlight(siteFrom, flightNumber, siteTo, aircraft);
            var isFlightPresent1 = flightMultiplModalpage.IsFirstFlightPresent();
            Assert.IsTrue(isFlightPresent1, "Aucune rule n'a été créée.");
            flightMultiplModalpage.RemoveFirstFlight();
            var isFlightPresent2 = flightMultiplModalpage.IsFirstFlightPresent();
            Assert.IsFalse(isFlightPresent2, "Le vol n'a pas été supprimé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Change_Flight()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);

            string firstFlightValue = flightPage.GetFirstFlightNumber();

            var editPage = flightPage.EditFirstFlight(flightNumber);
            string flightChanged = editPage.ChangeFlight(firstFlightValue);

            //Assert
            Assert.IsTrue(editPage.GetFlightNumber().Contains(flightChanged), "Le vol n'a pas été modifié.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Change_View()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.GetCalendarView();

            //Assert
            var verifyView = flightPage.IsViewChanged();
            Assert.IsTrue(verifyView, "La vue calendaire n'a pas été activée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_EditFlightType()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string flightNumberBis = flightNumber + "_bis";
            string typeValue = "Assistance";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            try
            {
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

                //Edit the flight type
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var flightEditPage = flightPage.EditFirstFlight(flightNumber);

                flightEditPage.EditFlightType(typeValue);

                flightPage = flightEditPage.CloseViewDetails();
                flightEditPage = flightPage.EditFirstFlight(flightNumber);
                var editedtype = flightEditPage.GetFlightType();

                //Assert
                Assert.AreEqual(editedtype, typeValue, "Les détails du vol n'ont pas été modifiés.");

            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, siteFrom, customer, false);
                flightPage.ResetFilter();
            }
            
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_CommentInternalFlightRemarks()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string flightNumberBis = flightNumber + "_bis";
            var flightRemarks = "XxxxX";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlightWithComment(flightNumber, customer, aircraft, siteFrom, siteTo);

            //Edit the first flight
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var flightEditPage = flightPage.EditFirstFlight(flightNumber);
            var editFlightInModal = flightEditPage.EditFlightInModal();
            editFlightInModal.SetFlightRemarks(flightRemarks);
            flightPage = flightEditPage.CloseViewDetails();

            flightPage.ModifyFlightName(flightNumberBis);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberBis);

            //Assert
            Assert.AreEqual(flightNumberBis.ToString(), flightPage.GetFirstFlightNumber(), "Les détails du vol n'ont pas été modifiés.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Edit_Flight()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string flightNumberBis = flightNumber + "_bis";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                //Edit the first flight
                flightPage.ModifyFlightName(flightNumberBis);
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberBis);
                //Assert 
                var displayedNumber = flightPage.GetFirstFlightNumber();
                Assert.AreEqual(flightNumberBis, displayedNumber, "Le vol n'a pas été édité.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumberBis);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Add_Guest_Type()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                var editFlight = flightPage.EditFirstFlight(flightNumber);
                editFlight.AddGuestType();
                var isGuestAdded = editFlight.IsGuestTypeVisible();
                Assert.IsTrue(isGuestAdded, "Le guest n'a pas été ajouté.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Replace_Guest_Type()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string initialGuest = "Leg 1 - FC";
            string newGuest = "SPML FC";
            string flightNumber = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            var editFlight = flightPage.EditFirstFlight(flightNumber);
            editFlight.AddGuestType();
            flightPage = editFlight.CloseViewDetails();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            editFlight = flightPage.EditFirstFlight(flightNumber);
            // le replace guest attéri dans l'index, mais faut refiltrer
            flightPage = editFlight.ReplaceGuest(initialGuest, newGuest);
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            editFlight = flightPage.EditFirstFlight(flightNumber);

            //Assert
            Assert.AreEqual(editFlight.GetGuestType(), newGuest, "Le guest n'a pas été ajouté.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Update_Guest_Qty()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string qtyGuest = "10";
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                flightPage.Filter(FilterType.SearchFlight, flightNumber);
                var editFlight = flightPage.EditFirstFlight(flightNumber);
                editFlight.AddGuestType();
                editFlight.UpdateGuestTypeQty(qtyGuest);
                flightPage = editFlight.CloseViewDetails();
                editFlight = flightPage.EditFirstFlight(flightNumber);
                //Assert
                var displayedQty = editFlight.GetGuestTypeQty();
                Assert.AreEqual(displayedQty, qtyGuest, "Le guest quantity n'a pas été ajouré.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Flight_History()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string qtyGuest = "10";
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                var editFlight = flightPage.EditFirstFlight(flightNumber);
                editFlight.AddGuestType();
                editFlight.UpdateGuestTypeQty(qtyGuest);
                flightPage = editFlight.CloseViewDetails();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                editFlight = flightPage.EditFirstFlight(flightNumber);
                //Assert
                var verifyHistory = editFlight.GetFlightHistory();
                Assert.IsTrue(verifyHistory, "L'historique du vol n'est pas visible.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Split_Flight()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString() + " to Split";

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

                var editFlight = flightPage.EditFirstFlight(flightNumber);
                editFlight.SplitFlight();
                editFlight.CloseViewDetails();

                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var firstValue = flightPage.GetFirstFlightNumber();

                flightPage.Filter(FlightPage.FilterType.Sites, siteTo);
                var secondValue = flightPage.GetFirstFlightNumber();
                var valuesCorrect = firstValue.Equals(secondValue);
                var arrowExist = flightPage.IsArrowSplitExist();
                //Assert
                Assert.IsTrue(valuesCorrect && arrowExist, "Le vol n'a pas été splité.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Add_Service()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            //Edit the first flight
            var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.AddGuestType();
            editPage.AddService(serviceName);

            //Assert
            Assert.AreEqual(1, editPage.CountServices(), MessageErreur.SERVICE_NON_AJOUTE);
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Delete_Service()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            //Edit the first flight
            var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.AddGuestType();
            editPage.AddService(serviceName);

            Assert.AreEqual(1, editPage.CountServices(), MessageErreur.SERVICE_NON_AJOUTE);

            flightPage = editPage.CloseViewDetails();
            flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.DeleteService();

            //Assert
            Assert.AreEqual(0, editPage.CountServices(), "Le service n'a pas été supprimé.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Import_Service()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string guestName = TestContext.Properties["Guest"].ToString();
            int qtyService = 4;
            string flightNumber = new Random().Next().ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            //Edit the first flight
            var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.AddGuestType(guestName);
            editPage.AddService(serviceName);
            flightPage = editPage.CloseViewDetails();
            editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.SetFinalQty(qtyService.ToString());
            flightPage = editPage.CloseViewDetails();

            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);

            flightPage.ClearDownloads();
            flightPage.ExportFlights(ExportType.ExportGuestAndServices, true);

            // On récupère les fichiers du répertoire de téléchargementf
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);
            flightPage.ClosePrintButton();
            editPage = flightPage.EditFirstFlight(flightNumber.ToString());

            editPage.DeleteGuestType();

            flightPage = editPage.ImportService(filePath);

            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            editPage = flightPage.EditFirstFlight(flightNumber.ToString());

            //Assert
            Assert.AreEqual(guestName, editPage.GetGuestType(), "Le leg n'a pas été importé.");
            Assert.AreEqual(qtyService.ToString(), editPage.GetServiceQty(), "La quantité n'a pas été importée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Duplicate_Flight()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string flightUpdate = flightNumber + "Updated";

            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                //Act
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

                //Edit the first flight
                var editPage = flightPage.EditFirstFlight(flightNumber);
                flightPage = editPage.DuplicateFlight(flightUpdate);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightUpdate);
                var firstFlightNumber = flightPage.GetFirstFlightNumber();
                //Assert
                Assert.AreEqual(flightUpdate, firstFlightNumber, "Le vol n'a pas été édité.");
               
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightUpdate);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Export_Fligthts_New_Version()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string flightNumber = new Random().Next().ToString();
            bool newVersion = true;

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            DeleteAllFileDownload();

            flightPage.ClearDownloads();

            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            //Edit the first flight
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType();
            editPage.AddService(serviceName);
            flightPage = editPage.CloseViewDetails();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            flightPage.ExportFlights(ExportType.Export, newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Flights", filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Export_Logs_Fligthts_New_Version()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string flightNumber = new Random().Next().ToString();
            bool newVersion = true;

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            flightPage.ClearDownloads();

            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            //Edit the first flight
            var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.AddGuestType();
            editPage.AddService(serviceName);
            flightPage = editPage.CloseViewDetails();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            flightPage.ExportFlights(ExportType.ExportLogs, newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("FlightsLogs", filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Export_OP245_Fligthts_New_Version()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string flightNumber = new Random().Next().ToString();
            bool newVersion = true;

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            flightPage.ClearDownloads();

            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            //Edit the first flight
            var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.AddGuestType();
            editPage.AddService(serviceName);
            flightPage = editPage.CloseViewDetails();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            flightPage.ExportFlights(ExportType.ExportOP, newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("MASTER", filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Delete_Guest_type()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            try
            {
                //Edit the first flight
                var editPage = flightPage.EditFirstFlight(flightNumber);
                editPage.AddGuestType();
                editPage.DeleteGuestType();
                bool isGuestVisible = editPage.IsGuestTypeVisible();
                //Assert
                Assert.IsFalse(isGuestVisible, "Le guest type n'a pas été supprimé.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();
                massiveDeleteModal.SetFlightName(flightNumber);
                massiveDeleteModal.ClickSearchButton();
                massiveDeleteModal.ClickSelectAllButton();
                massiveDeleteModal.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Test_Flight_Service_Generique_Add()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                var editPage = flightPage.EditFirstFlight(flightNumber);
                editPage.AddGuestType();
                editPage.AddGenericService();
                //Assert
                var isServiceAdded = editPage.IsGenericServiceAdded();
                Assert.IsTrue(isServiceAdded, "Le service générique n'a pas été ajouté.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Test_Flight_Service_Generique_Create()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
                var serviceNameTest = "ServiceNameTest" + new Random().Next().ToString();
                editPage.AddGuestType();
                editPage.CreateGenericService(serviceNameTest);
                //Assert
                var isAdded = editPage.IsGenericServiceAdded();
                Assert.IsTrue(isAdded, "Le service n'a pas été créé.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Print_CheckList_Tests_New_Version()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var customer = TestContext.Properties["CustomerLP"].ToString();
            var flightNumber = new Random().Next().ToString();
            bool versionprint = true;
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Print CheckList";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.ClearDownloads();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                FlightDetailsPage edit = flightPage.EditFirstFlight(flightNumber);
                edit.AddGuestType("YC");
                edit.AddService("Service for flight");
                flightPage = edit.CloseViewDetails();
                flightPage.SetNewState("V");

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                // Print check list
                var reportPage = flightPage.PrintReport(PrintType.Checklist, versionprint);
                bool isPrintCheckList = reportPage.IsReportGenerated();
                reportPage.Close();
                //Assert        
                Assert.IsTrue(isPrintCheckList, "Le checklist n'a pas été générée.");
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                reportPage.Close();
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fichier = new FileInfo(trouve);
                fichier.Refresh();
                Assert.IsTrue(fichier.Exists, "Pas de print pdf");
                //Le PDF imprimé devrait contenir:
                PdfDocument document = PdfDocument.Open(fichier.FullName);
                string mots = "";
                foreach (Page p in document.GetPages())
                {
                    // pdf Iron découpe les mots
                    mots += p.Text.Replace(" ", "");
                }
                var filtredData = mots.Contains(siteFrom);
                Assert.IsTrue(filtredData, MessageErreur.EXCEL_DONNEES_KO);
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Print_Turnover_Tests_New_Version()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var customer = TestContext.Properties["CustomerLP"].ToString();

            var flightNumber = new Random().Next();
            bool versionprint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            flightPage.ClearDownloads();

            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber.ToString(), customer, aircraft, siteFrom, siteTo);

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            // Print turnover
            var reportPage = flightPage.PrintReport(PrintType.Turnover, versionprint);
            bool isPrintTurnover = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isPrintTurnover, "Le turnover n'a pas été généré.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Print_TrolleysLabels()
        {
            // Prepare
            bool versionprint = true;
            string site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["Aircraft"].ToString();
            var route = "ACE-BCN";
            string customer = "$$ - CAT Genérico";
            DateTime fromDate = DateUtils.Now.AddMonths(-1);
            DateTime toDate = DateUtils.Now.AddMonths(3);
            DateTime flightDate = DateUtils.Now;
            string lpCartName = "lpcart" + "-" + new Random().Next().ToString();
            string serviceName = "serviceNameToday" + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string loadingPlanName = "Lodplan " + "-" + new Random().Next().ToString();
            string guestName = "FC";
            string code = "codeLpcart" + "-" + new Random().Next().ToString();
            string comment = "Bob comment";
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string etaHours = "00";
            string etdHours = "23";
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();
            string datasheet = TestContext.Properties["DatasheetName"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Print trolleys";
            string position = "accotroScheme";
            string qty = "2";
            string rows = "2";
            string columns = "2";
            var deliveryNote = "New2";

            //Arrange
            HomePage homePage = LogInAsAdmin();
            var paramPage = homePage.GoToParameters_GlobalSettings();
            var paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();

            var flightPage = new FlightPage(WebDriver, TestContext);
            flightPage.RemoveAllFromSelectedDeliveryNoteReport();
            #region refresh
            homePage.Navigate();
            paramPage = homePage.GoToParameters_GlobalSettings();
            paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            #endregion

            flightPage.SelectDeliveryNoteReport(deliveryNote);

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
            pricePage.UnfoldAll();
            pricePage.SetFirstPriceDataSheet(datasheet);
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
            loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customer, route, aircraft, site);
            loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(toDate, fromDate);
            var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
            loadingPlanDetailsPage.ClickAddGuestBtn();
            loadingPlanDetailsPage.SelectGuest(guestName);
            loadingPlanDetailsPage.ClickCreateGuestBtn();
            loadingPlanDetailsPage.ClickFirstGuest();
            loadingPlanDetailsPage.AddServiceBtn();
            loadingPlanDetailsPage.AddNewService(serviceName);
            servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            pricePage = servicePage.ClickOnFirstService();
            var serviceLoadingPage = pricePage.GoToLoadingPlanTab();
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartCreateModalPage = lpCartPage.LpCartCreatePage();
            var LpCartDetailPage = lpCartCreateModalPage.FillField_CreateNewLpCartWithRoutes(code, lpCartName, site, customer, aircraft, fromDate, toDate, comment, route);
            LpCartDetailPage.ClickAddtrolley();
            LpCartDetailPage.AddTrolley(trolleyName);
            var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
            lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, rows, columns);
            LpCartDetailPage.BackToList();
            flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FilterType.Sites, siteFrom);
            var flightCreatePageModal = flightPage.FlightCreatePage();
            flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, flightDate);
            try
            {
                flightCreatePageModal.ClearDownloads();
                var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();
                trolleyLightLabelPage.ResetFilters();
                var trolleyDetailledPage = trolleyLightLabelPage.GoToDetailledLabel();
                trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.SearchFlight, flightNumber);
                trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.Site, site);
                trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));
                trolleyDetailledPage.AddPositionAndQuantity(position, qty);
                flightPage = homePage.GoToFlights_FlightPage();
                // Print Trolleyslabels
                DeleteAllFileDownload();
                flightPage.ClearDownloads();
                var reportPage = flightPage.PrintReport(PrintType.Trolleyslabels, versionprint);
                reportPage.Close();
                PrintReportPage printReport = flightPage.ClickPrinterIcon();
                printReport.ClickRefreshUntilFinished();
                printReport.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string pdfFoundPath = printReport.PrintAllZipAllPdf(downloadsPath, DocFileNameZipBegin, false);
                PdfKeywordSearcher searcher = new PdfKeywordSearcher();
                bool pdfsFound = searcher.FindAllPdfsWithKeywordMeal(downloadsPath, trolleyScheme);
                // Assert 
                Assert.IsTrue(pdfsFound, $"No PDF containing trolleyslabels was found in the directory: {downloadsPath}");
            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
                //Delete LoadingPlan
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, null, null, toDate.AddMonths(2));
                //Delete LpCart
                lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                lpCartPage.DeleteLpCart();
                //Delete service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Print_LpCartSeals()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var customer = TestContext.Properties["CustomerLP"].ToString();
            var flightNumber = new Random().Next().ToString();
            bool versionprint = true;

            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.ClearDownloads();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                // Print LpCartSeals
                var reportPage = flightPage.PrintReport(PrintType.LpCartSeals, versionprint);
                bool isPrintLpCartSeals = reportPage.IsReportGenerated();
                reportPage.Close();
                //Assert
                Assert.IsTrue(isPrintLpCartSeals, "LpCartSeals n'a pas été généré.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Print_SpecialMeals()
        {
            // Prepare
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var flightNumber = new Random().Next().ToString();
            bool versionprint = true;
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string specialmeals = "SPML";
            string guestType = "SPML BC";
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Print Special";
            bool withRefresh = false;
            string route = "ACE-BCN";

            var homePage = LogInAsAdmin();
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            serviceCreateModalPage.SetSPML(specialmeals);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(siteFrom, customer, fromDate, toDate);
            pricePage.ToggleFirstPrice();
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var editPage = flightPage.EditFirstFlight(flightNumber);
                editPage.AddGuestTypeMeal(route, guestType);
                editPage.AddService(serviceName);
                homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                DeleteAllFileDownload();
                flightPage.ClearDownloads();
                var reportPage = flightPage.PrintReport(PrintType.SpecialMeals, versionprint);
                reportPage.Close();
                PrintReportPage printReport = flightPage.ClickPrinterIcon();
                printReport.ClickRefreshUntilFinished();
                printReport.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string pdfFoundPath = printReport.PrintAllZipAllPdf(downloadsPath, DocFileNameZipBegin, withRefresh);
                PdfKeywordSearcher searcher = new PdfKeywordSearcher();
                FileInfo filePdf = new FileInfo(pdfFoundPath);
                filePdf.Refresh();
                Assert.IsTrue(filePdf.Exists, pdfFoundPath + " non généré");
                //Le PDF imprimé devrait contenir:
                PdfDocument document = PdfDocument.Open(filePdf.FullName);
                List<string> mots = new List<string>();
                foreach (Page p in document.GetPages())
                {
                    foreach (var mot in p.GetWords())
                    {
                        mots.Add(mot.Text);
                    }
                }
                var verifFlightNumber = mots.Contains(flightNumber);
                var verifCompare = mots.Contains(specialmeals);
                var verifData = verifFlightNumber && verifCompare;
                // Assert 
                Assert.IsTrue(verifData, $"No PDF containing specialmeals was found in the directory: {downloadsPath}");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Print_SwapLpCartReport()
        {
            // Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["Aircraft"].ToString();
            var route = "ACE-BCN";
            string customer = "$$ - CAT Genérico";
            DateTime fromDate = DateUtils.Now.AddDays(-1);
            DateTime toDate = DateUtils.Now.AddMonths(3);
            string lpCartName = "lpcart" + "-" + new Random().Next().ToString();
            string serviceName = "serviceNameToday" + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string loadingPlanName = "Lodplan " + "-" + new Random().Next().ToString();
            string guestName = "FC";
            string code = "codeLpcart" + "-" + new Random().Next().ToString();
            string comment = "Bob comment";
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string etaHours = "00";
            string etdHours = "23";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Flights Swap LPCart Report";
            string DocFileNameZipBegin = "All_files_";
            var compareLpcart = "SAME";

            //Arrange
            HomePage homePage = LogInAsAdmin();
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
            loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customer, route, aircraft, site);
            loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(toDate, fromDate);
            var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
            loadingPlanDetailsPage.ClickAddGuestBtn();
            loadingPlanDetailsPage.SelectGuest(guestName);
            loadingPlanDetailsPage.ClickCreateGuestBtn();
            loadingPlanDetailsPage.ClickFirstGuest();
            loadingPlanDetailsPage.AddServiceBtn();
            loadingPlanDetailsPage.AddNewService(serviceName);
            servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            pricePage = servicePage.ClickOnFirstService();
            var serviceLoadingPage = pricePage.GoToLoadingPlanTab();
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartCreateModalPage = lpCartPage.LpCartCreatePage();
            var LpCartDetailPage = lpCartCreateModalPage.FillField_CreateNewLpCartWithRoutes(code, lpCartName, site, customer, aircraft, fromDate, toDate, comment, route);
            LpCartDetailPage.ClickAddtrolley();
            LpCartDetailPage.AddTrolley(trolleyName);
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter(); // ACE from
            var flightCreatePageModal = flightPage.FlightCreatePage();
            flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, fromDate);
            try
            {
                flightCreatePageModal.ClearDownloads();
                FlightDetailsPage FlightDetailsPage = flightPage.EditFirstFlight(flightNumber);
                FlightDetailsPage.ShowInfoFlightExtendedMenu();

                // Print
                PrintReportPage PrintReportPage = FlightDetailsPage.SwapLPCartReport();
                bool isPrintSwapLpCartReport = PrintReportPage.IsReportGenerated();
                PrintReportPage.Close();
                //Assert
                Assert.IsTrue(isPrintSwapLpCartReport, "SwapLpCartReport  n'a pas été généré.");

                PrintReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                // cliquer sur All
                string trouve = PrintReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
                FileInfo filePdf = new FileInfo(trouve);
                filePdf.Refresh();
                Assert.IsTrue(filePdf.Exists, trouve + " non généré");

                //Le PDF imprimé devrait contenir:
                PdfDocument document = PdfDocument.Open(filePdf.FullName);
                List<string> mots = new List<string>();
                foreach (Page p in document.GetPages())
                {
                    foreach (var mot in p.GetWords())
                    {
                        mots.Add(mot.Text);
                    }
                }
                var verifFlightNumber = mots.Contains(flightNumber);
                var verifLpCart = mots.Count(x => x.Equals(lpCartName)) == 2;
                var verifCompare = mots.Contains(compareLpcart);
                var verifData = verifFlightNumber && verifLpCart && verifCompare;
                Assert.IsTrue(verifData, "data in pdf sont pas corrects");
            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
                //Delete LoadingPlan
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, null, null, toDate.AddMonths(2));
                //Delete LpCart
                lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                lpCartPage.DeleteLpCart();
                //Delete service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Print_TrackSheet_Tests_New_Version()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var customer = TestContext.Properties["CustomerLP"].ToString();

            var flightNumber = new Random().Next();
            bool versionprint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            flightPage.ClearDownloads();

            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber.ToString(), customer, aircraft, siteFrom, siteTo);

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            // Print track sheet
            var reportPage = flightPage.PrintReport(PrintType.TrackSheet, versionprint);
            bool isPrintTrackSheet = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isPrintTrackSheet, "Le trackshhet n'a pas été généré.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Print_DeliveryNote_Tests_New_Version()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            bool versionprint = true;

            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.ClearDownloads();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            // Print delivery note
            var reportPage = flightPage.PrintReport(PrintType.DeliveryNote, versionprint);
            bool isPrintDeveliryNote = reportPage.IsReportGenerated();
            reportPage.Close();
            //Assert
            Assert.IsTrue(isPrintDeveliryNote, "Le rapport n'a pas été généré.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Print_DeliveryBOB_New_Version()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var customer = TestContext.Properties["CustomerLP"].ToString();
            var flightNumber = new Random().Next().ToString();
            bool versionprint = true;

            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.ClearDownloads();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                // Print delivery note
                var reportPage = flightPage.PrintReport(PrintType.DeliveryBOB, versionprint);
                bool isPrintDeveliryNote = reportPage.IsReportGenerated();
                reportPage.Close();
                //Assert
                Assert.IsTrue(isPrintDeveliryNote, "Le rapport n'a pas été généré.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Change_Date_View()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var customer = TestContext.Properties["CustomerLP"].ToString();

            var flightNumber = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            // Change date (get the last date)
            flightPage.SetDateState(DateUtils.Now.AddDays(-1));
            var stateDate = flightPage.GetDateState();

            if (flightPage.CheckTotalNumber() == 0)
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber.ToString(), customer, aircraft, siteFrom, siteTo);
            }
            else
            {
                flightNumber = flightPage.GetFirstFlightNumber();
            }

            // Get date for the first line in detail view
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber.ToString());
            var flightEditPage = flightPage.EditFirstFlight(flightNumber.ToString());

            // Assert
            Assert.AreEqual(flightEditPage.GetDateConfigured(), stateDate.Split(' ')[0], "La date n'a pas été modifiée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Filter_Display()
        {
            var displayFlightDeparture = "Display flight departure";
            var displayFlightArrival = "Display flight arrival";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            var site = flightPage.GetSiteFilter();
            flightPage.Filter(FilterType.DispalyFlight, displayFlightDeparture);
            var from = flightPage.VerifySiteFrom(site);
            flightPage.Filter(FilterType.DispalyFlight, displayFlightArrival);
            var to = flightPage.VerifySiteTo(site);
            Assert.IsTrue(from, "erreur de filtrage");
            Assert.IsTrue(to, "erreur de filtrage");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Index_PrintFlightsAssignment()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Flights assignment report";
            string DocFileNameZipBegin = "All_files_";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, "ACE");
            var flightNo = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNo);
            flightPage.ClearDownloads();
            var reportPage = flightPage.PrintReport(PrintType.FlightAssignments, true);
            bool isPrintDeveliryNote = reportPage.IsReportGenerated();
            reportPage.Close();
            Assert.IsTrue(isPrintDeveliryNote, "Le rapport n'a pas été généré.");

            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");
            Assert.IsTrue(flightPage.VerifierFlightNumberInPdf(filePdf, flightNo), "data in pdf sont pas corrects");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Index_PrintFlightsTempChecksList()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Flights TempChecks report";
            string DocFileNameZipBegin = "All_files_";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, "ACE");
            var flightNo = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNo);
            flightPage.ClearDownloads();
            var reportPage = flightPage.PrintReport(PrintType.FlightsTempChecksList, true);
            bool isPrinted = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isPrinted, "Le rapport n'a pas été généré.");

            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");
            Assert.IsTrue(flightPage.VerifierFlightNumberInPdf(filePdf, flightNo), "data in pdf sont pas corrects");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_RefreshInBoundInfo()
        {
            //Prepare

            //Auth
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, "ACE");
            string flightNumber = flightPage.GetFirstFlightNumber();
            flightPage.RefreshInBoundInfo(flightNumber);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Details_PrintFlightLabels()
        {
            //Prepare
            string name = "lpcart_For_Flight_Label";
            string code = "Labellp";
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = "TVS - SMARTWINGS, A.S. (TVS)";
            string aircraft = "AB318";
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Label comment";

            string trolley = "TrolleyL";

            var flightNo = new Random().Next().ToString();
            string flightType = "Regular";
            var siteFrom = TestContext.Properties["SiteLpCart"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();

            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string DocFileNamePdfBegin = "Flights Print Label";
            string DocFileNameZipBegin = "All_files_";

            //Auth
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.ClearDownloads();
            string dateFormat = homePage.GetDateFormatPickerValue();
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            //Act
            //Print Flight Label depuis le detail

            //Choisir le format du label dans Customer : square ou [portrait]
            CustomerPage customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.ResetFilters();
            customerPage.Filter(CustomerPage.FilterType.Search, "TVS");
            customerPage.Filter(CustomerPage.FilterType.SortBy, "CUSTOMER");
            CustomerGeneralInformationPage customerGeneralInfo = customerPage.SelectFirstCustomer();
            //voir FL_FLIG_Index_PrintFlightLabels, pour le test Portrait.
            customerGeneralInfo.SetPrintLabelFormat("Square");

            //Avoir des LP Carts crées dans le module LP Carts, avec des Trolleys à l'intérieur (Carts).
            //1.Entrer sur LP Carts
            LpCartPage lpCart = homePage.GoToFlights_LpCartPage();
            lpCart.ResetFilter();
            lpCart.Filter(LpCartPage.FilterType.Search, code);
            if (lpCart.CheckTotalNumber() == 1)
            {
                // use case
                LpCartCartDetailPage lapin = lpCart.ClickFirstLpCart();
                LpCartFlightDetailPage flightTab = lapin.ClickOnFlightTab();
                List<string> names = flightTab.GetFlightName();
                List<string> dates = flightTab.GetFlightDates();
                FlightPage flightPage = flightTab.GoToFlights_FlightPage();
                //retirer les lp des flights du use case
                for (int i = 0; i < names.Count; i++)
                {
                    flightPage.ResetFilter();
                    flightPage.SetDateState(DateTime.Parse(dates[i], ci));
                    flightPage.Filter(FlightPage.FilterType.Sites, site);
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, names[i]);
                    if (flightPage.IsValidated())
                    {
                        flightPage.UnSetNewState("V");
                    }
                    if (flightPage.IsFlightExist())
                    {
                        var flightModalEdit = flightPage.EditFirstFlight(names[i]);
                        flightModalEdit.SelectLpCart("None");
                        flightPage = flightModalEdit.CloseViewDetails();
                    }
                }
                lpCart = flightPage.GoToFlights_LpCartPage();
                lpCart.ResetFilter();
                lpCart.Filter(LpCartPage.FilterType.Search, code);
                lpCart.DeleteLpCart();
            }
            //2.Créer un LP Cart[+] - Create LP Cart
            LpCartCreateModalPage modal = lpCart.LpCartCreatePage();
            //3.Ajouter un Code, Name, Site (Houston), Customer, Aircraft Type, From - To, Routes(taper IAH - SFO, add, IAH - ABQ, IAH - MAF...)
            LpCartCartDetailPage details = modal.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);
            LpCartGeneralInformationPage generalInfo = details.LpCartGeneralInformationPage();
            generalInfo.AddRoute("ACE - AGP");
            generalInfo.AddRoute("ACE - BCN");
            generalInfo.AddRoute("ACE - BFS");
            //details
            details = generalInfo.ClickOnCarts();
            //4.Save

            //Créer des Carts
            //5.Cliquer sur le LP Cart créé,
            //6.Cliquer sur[+](d'Add)
            //7.Entrer le Trolley Name, Entrer le Trolley Code, Entrer la Location(se baser sur une existant ex.DT, DL, DP, DF), Galley Code(de G1 à G4), Galley Location(United Economy, United First), Equipment.
            details.ClickAddtrolley();
            details.AddTrolley(trolley, "shortC");
            //8.Entrer dans le cart(stylo d'Edit Cart)
            LpCartEditLpCartSchemeModal modalCartScheme = details.EditCartScheme();
            //9.Ajouter des colonnes, détails, enregistrer
            modalCartScheme.EditLpCartscheme(trolley, "2", "2", "2", "2", "");

            //Ajouter mon LPCart à mes vols
            //10.Aller sur Flights
            FlightPage flights = details.GoToFlights_FlightPage();
            //11.Choisir une date(qui est cohérente avec ma création des LP Carts dessus)
            flights.WaiForLoad();
            FlightCreateModalPage modalFlight = flights.FlightCreatePage();
            modalFlight.FillField_CreatNewFlightWithFlightType(flightNo, customer, aircraft, siteFrom, siteTo, flightType, null, "00", "23", null, from);
            homePage.Navigate();
            //12.Sur le résultats de vols sur l'index, ajouter vol par vol un LP Cart
            flights = homePage.GoToFlights_FlightPage();
            flights.ResetFilter();
            flights.SetDateState(from);
            flights.Filter(FilterType.SearchFlight, flightNo);
            flights.Filter(FilterType.Sites, site);
            //12.1 Click sur le vol picto de tout à droit
            FlightDetailsPage detailsFlight = flights.EditFirstFlight(flightNo);
            detailsFlight.SetRegistration("reg_999");
            //12.2 Sur la colonne de Gauche clicker sur le menu déroulant "LP Cart"
            //12.3 Ajouter un LP Cart
            // undo via Registration
            detailsFlight.SelectLpCart(name);
            //13.Attendre d'avoir le refresh de mon Flight Detail
            //14.Sortir du Flight Detail
            flights = detailsFlight.CloseViewDetails();
            flights.CliCkOnFirstFlight(2);
            // undo via Registration
            flights.SetDockNumber("777");
            //Refaire sur chaque vol

            //15.Sur l'index flights, ajouter le Dock Number
            flights.ResetFilter();
            flights.SetDateState(from);
            flights.Filter(FilterType.SearchFlight, flightNo);
            flights.Filter(FilterType.Sites, site);

            //1.Entrer sur le programme Flights > Flights >
            //----> SI j'ai plusieurs vols dans mon index, Choisir un EDT From et EDT To qui pourrait réduire le nombre de vols
            //2.Aller sur le detail d'un Flight
            detailsFlight = flights.EditFirstFlight(flightNo);
            //3.Cliquer... et sur "Print Flight Labels"
            PrintReportPage printPage = detailsFlight.PrintFlightLabels();
            printPage.IsReportGenerated();
            printPage.Close();
            //download pdf file
            printPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = printPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fichier = new FileInfo(trouve);
            fichier.Refresh();
            Assert.IsTrue(fichier.Exists, "Pas de print pdf");

            //4.Ajouter un Workshop, laisser en blanc le champ
            TimeBlockPage tb = detailsFlight.GoToFlight_TimeBlockPage();
            tb.ResetFilters();
            tb.Filter(TimeBlockPage.FilterType.DateFrom, from);
            tb.Filter(TimeBlockPage.FilterType.DateTo, to);
            //tb.WaitPageLoading();
            tb.Filter(TimeBlockPage.FilterType.Search, flightNo);
            tb.WaitLoading();
            Assert.AreEqual(1, tb.CheckTotalNumber(), "Pas de Workshop");

            //5.Imprimer

            //Le PDF imprimé devrait contenir:
            PdfDocument document = PdfDocument.Open(fichier.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }
            //1er tableau
            //-1ère ligne: flight date +customer / flight nb + trolley preparation zone
            Assert.AreEqual(2, mots.Count(x => x == from.ToString("MMM", CultureInfo.GetCultureInfo("en-US"))), "pas de date flight dans le pdf");
            Assert.AreEqual(2, mots.Count(x => x == "TVS"), "pas de customer dans le pdf");
            Assert.AreEqual(2, mots.Count(x => flightNo.StartsWith(x) && x.Length >= 7), "pas de flight nb dans le pdf");
            // pas 14 Trolley car certain TrolleyL sont fusionné dans le PDF
            Assert.AreEqual(11, mots.Count(x => (x.Contains("Trolley") || "Trolley".StartsWith(x)) && x.Length >= 5), "pas de trolley label dans le pdf (1)");
            //-2e ligne: route, aircraft, registration, menu code(airvision) +dock nb + Workhops + ETD + Trolley Name
            Assert.AreEqual(2, mots.Count(x => x == "ACE"), "pas de route dans le pdf (1)");
            Assert.AreEqual(2, mots.Count(x => x == "AGP"), "pas de route dans le pdf (2)");
            // registration : pars à la dérive
            Assert.AreEqual(2, mots.Count(x => x == "Ensambl"), "pas de Workflow Ensambl dans le pdf");
            Assert.AreEqual(2, mots.Count(x => x == "Corte"), "pas de Workflow Corte dans le pdf");
            Assert.AreEqual(6, mots.Count(x => x == "23:00"), "pas de ETD dans le pdf (2)");
            // Trolley Label déjà compté

            //2ème tableau
            //-1ère ligne : galley location +aircraft + start date LPcart
            // Trolley Label déjà compté 3 fois
            Assert.AreEqual(2, mots.Count(x => aircraft.StartsWith(x) && x.Length >= 4), "pas de aircraft/registration dans le pdf");
            Assert.AreEqual(2, mots.Count(x => x == from.ToString("MMM", CultureInfo.GetCultureInfo("en-US"))), "pas de start date LPcart dans le pdf");
            //-2ème ligne : trolley name +galley code + LP cart comment
            //-comment : trolley comment
            // Trolley Label déjà compté 1 fois

            //3ème tableau
            //Front and Back(trolley scheme detail)
            Assert.AreEqual(4, mots.Count(x => x == "A00:" || x == "B00"), "pas de Front and Back dans le pdf (1)");
            Assert.AreEqual(4, mots.Count(x => x == "A10:" || x == "B10"), "pas de Front and Back dans le pdf (2)");
            Assert.AreEqual(4, mots.Count(x => x == "A01:" || x == "B01"), "pas de Front and Back dans le pdf (3)");
            Assert.AreEqual(4, mots.Count(x => x == "A11:" || x == "B11"), "pas de Front and Back dans le pdf (4)");
            //Trolley equipment -Trolley code
            // Trolley Label déjà compté 2 fois
            //LP cart name
            Assert.AreEqual(2, mots.Count(x => x == name), "pas de LP cart name dans le pdf");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Index_PrintFlightLabels()
        {
            //Print Flight Label depuis l'index de Flights

            //Prepare
            string name = "lpcart_For_Flight_Labels";
            string code = "Labellp";
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = "TVS - SMARTWINGS, A.S. (TVS)";
            string aircraft = "AB318";
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Label comment";

            string trolley = "Trolley Label";

            var flightNo = new Random().Next().ToString();
            string flightType = "Regular";
            var siteFrom = TestContext.Properties["SiteLpCart"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();

            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string DocFileNamePdfBegin = "Flights Print Label";
            string DocFileNameZipBegin = "All_files_";
            //Auth
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.ClearDownloads();
            string dateFormat = homePage.GetDateFormatPickerValue();
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            //Act
            //Print Flight Label depuis l'index de Flights

            //Choisir le format du label dans Customer : square ou [portrait]
            CustomerPage customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.ResetFilters();
            customerPage.Filter(CustomerPage.FilterType.Search, "TVS");
            customerPage.Filter(CustomerPage.FilterType.SortBy, "CUSTOMER");
            CustomerGeneralInformationPage customerGeneralInfo = customerPage.SelectFirstCustomer();
            //voir FL_FLIG_Details_PrintFlightLabels, pour le test Square.
            customerGeneralInfo.SetPrintLabelFormat("Portrait");

            //Parmaétrer les workshops: Dans Parameters > Tablet : cocher on Show on print label(ici Ensamblaje et Corte)

            //Avoir des LP Carts crées dans le module LP Carts, avec des Trolleys à l'intérieur (Carts).
            //1.Entrer sur LP Carts
            LpCartPage lpCart = homePage.GoToFlights_LpCartPage();
            lpCart.ResetFilter();
            lpCart.Filter(LpCartPage.FilterType.Search, code);
            if (lpCart.CheckTotalNumber() == 1)
            {
                // use case
                LpCartCartDetailPage lapin = lpCart.ClickFirstLpCart();
                LpCartFlightDetailPage flightTab = lapin.ClickOnFlightTab();
                List<string> names = flightTab.GetFlightName();
                List<string> dates = flightTab.GetFlightDates();
                FlightPage flightPage = flightTab.GoToFlights_FlightPage();
                //retirer les lp des flights du use case
                for (int i = 0; i < names.Count; i++)
                {
                    flightPage.ResetFilter();
                    flightPage.SetDateState(DateTime.Parse(dates[i], ci));
                    flightPage.Filter(FlightPage.FilterType.Sites, site);
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, names[i]);
                    if (flightPage.IsValidated())
                    {
                        flightPage.UnSetNewState("V");
                    }
                    if (flightPage.IsFlightExist())
                    {
                        var flightModalEdit = flightPage.EditFirstFlight(names[i]);
                        flightModalEdit.SelectLpCart("None");
                        flightPage = flightModalEdit.CloseViewDetails();
                    }
                }
                lpCart = flightPage.GoToFlights_LpCartPage();
                lpCart.ResetFilter();
                lpCart.Filter(LpCartPage.FilterType.Search, code);
                lpCart.DeleteLpCart();
            }
            //2.Créer un LP Cart[+] - Create LP Cart
            LpCartCreateModalPage modal = lpCart.LpCartCreatePage();
            //3.Ajouter un Code, Name, Site (Houston), Customer, Aircraft Type, From - To, Routes(taper IAH - SFO, add, IAH - ABQ, IAH - MAF...)
            LpCartCartDetailPage details = modal.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);
            LpCartGeneralInformationPage generalInfo = details.LpCartGeneralInformationPage();
            generalInfo.AddRoute("ACE - AGP");
            generalInfo.AddRoute("ACE - BCN");
            generalInfo.AddRoute("ACE - BFS");
            //details
            details = generalInfo.ClickOnCarts();
            //4.Save

            //Créer des Carts
            //5.Cliquer sur le LP Cart créé,
            //6.Cliquer sur[+](d'Add)
            //7.Entrer le Trolley Name, Entrer le Trolley Code, Entrer la Location(se baser sur une existant ex.DT, DL, DP, DF), Galley Code(de G1 à G4), Galley Location(United Economy, United First), Equipment.
            details.ClickAddtrolley();
            details.AddTrolley(trolley, "shortC");
            //8.Entrer dans le cart(stylo d'Edit Cart)
            LpCartEditLpCartSchemeModal modalCartScheme = details.EditCartScheme();
            //9.Ajouter des tiroirs, détails, enregistrer
            modalCartScheme.EditLpCartscheme(trolley, "2", "2", "2", "2", "");

            //Ajouter mon LPCart à mes vols
            //10.Aller sur Flights
            FlightPage flights = details.GoToFlights_FlightPage();
            flights.ResetFilter();
            flights.Filter(FlightPage.FilterType.Sites, siteFrom);
            //11.Choisir une date(qui est cohérente avec ma création des LP Carts dessus)
            FlightCreateModalPage modalFlight = flights.FlightCreatePage();
            modalFlight.FillField_CreatNewFlightWithFlightType(flightNo, customer, aircraft, siteFrom, siteTo, flightType, null, "00", "23", null, from);

            //homePage.Navigate();
            //12.Sur le résultats de vols sur l'index, ajouter vol par vol un LP Cart
            //flights = homePage.GoToFlights_FlightPage();
            flights.ResetFilter();
            flights.SetDateState(from);
            flights.Filter(FilterType.Sites, siteFrom);
            flights.Filter(FilterType.SearchFlight, flightNo);
            Assert.AreEqual(1, flights.CheckTotalNumber(), " Vol " + flightNo + " non présent à la date " + from.ToString("yyyy-MM-dd") + " cas 1");
            flights.Filter(FilterType.Sites, site);
            Assert.AreEqual(1, flights.CheckTotalNumber(), " Vol " + flightNo + " non présent à la date " + from.ToString("yyyy-MM-dd") + " cas 2");

            //12.1 Click sur le vol picto de tout à droit
            FlightDetailsPage detailsFlight = flights.EditFirstFlight(flightNo);
            detailsFlight.SetRegistration("reg_999");
            //Attention le vol doit avoir un guest et un service
            detailsFlight.AddGuestType("FC");
            detailsFlight.AddService("AIR FRESHNER TVS");
            //detailsFlight.
            //12.2 Sur la colonne de Gauche clicker sur le menu déroulant "LP Cart"
            //12.3 Ajouter un LP Cart
            // undo via Registration
            detailsFlight.SelectLpCart(name);
            //13.Attendre d'avoir le refresh de mon Flight Detail
            //14.Sortir du Flight Detail
            flights = detailsFlight.CloseViewDetails();
            flights.CliCkOnFirstFlight(2);
            //Refaire sur chaque vol

            // rafraichir le Regist. REG_999
            flights.ResetFilter();
            flights.SetDateState(from);
            flights.Filter(FilterType.Sites, siteFrom);
            flights.Filter(FilterType.SearchFlight, flightNo);
            Assert.AreEqual(1, flights.CheckTotalNumber(), " Vol " + flightNo + " non présent à la date " + from.ToString("yyyy-MM-dd") + " cas 1");
            flights.Filter(FilterType.Sites, site);
            Assert.AreEqual(1, flights.CheckTotalNumber(), " Vol " + flightNo + " non présent à la date " + from.ToString("yyyy-MM-dd") + " cas 2");
            flights.CliCkOnFirstFlight(2);

            //15.Sur l'index flights, ajouter le Dock Number
            // undo via Registration
            flights.SetDockNumber("777");

            //1.Entrer sur le programme Flights > Flights >
            //----> SI j'ai plusieurs vols dans mon index, Choisir un EDT From et EDT To qui pourrait réduire le nombre de vols
            //2.Aller sur[...] en haut à droit
            //3.Cliquer sur "Print Flight Labels"
            flights.ClearDownloads();
            flights.PurgeDownloads();
            PrintReportPage reportPage = flights.PrintReport(PrintType.FlightsLabels, true);
            //4.Ajouter un Workshop, laisser en blanc le champ
            //5.Imprimer
            bool isPrintFlightLabels = reportPage.IsReportGenerated();
            reportPage.Close();
            Assert.IsTrue(isPrintFlightLabels, "Le rapport n'a pas été généré.");
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            //Le PDF imprimé devrait contenir:
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }

            //1er tableau
            //-1ère ligne: flight date +customer / flight nb + trolley preparation zone
            Assert.AreEqual(2, mots.Count(x => x == from.ToString("MMM", CultureInfo.GetCultureInfo("en-US"))), "pas de date flight dans le pdf");
            Assert.AreEqual(2, mots.Count(x => x == "TVS"), "pas de customer dans le pdf");
            Assert.AreEqual(2, mots.Count(x => x == flightNo), "pas de flight nb dans le pdf");
            Assert.AreEqual(2 + 2 + 6 + 2 + 4, mots.Count(x => x == "Trolley"), "pas de trolley label dans le pdf (1)");
            Assert.AreEqual(2 + 5 + 2 + 2 + 5 + 2, mots.Count(x => x == "Label"), "pas de trolley label dans le pdf (2)");  // dont le nom de la page tout en bas
                                                                                                                            //-2e ligne: route, aircraft, registration, menu code(airvision) +dock nb + Workhops + ETD + Trolley Name
            Assert.AreEqual(2, mots.Count(x => x == "ACE"), "pas de route dans le pdf (1)");
            Assert.AreEqual(2, mots.Count(x => x == "AGP"), "pas de route dans le pdf (2)");
            Assert.AreEqual(2, mots.Count(x => x == aircraft + "-REG_999"), "pas de aircraft/registration dans le pdf");
            // registration : pars à la dérive
            Assert.AreEqual(2, mots.Count(x => x == "Dock:"), "pas de Dock dans le pdf (1)");
            Assert.AreEqual(2, mots.Count(x => x == "777"), "pas de Dock dans le pdf (2)");
            Assert.AreEqual(2, mots.Count(x => x == "Ensambl"), "pas de Workflow Ensambl dans le pdf");
            Assert.AreEqual(2, mots.Count(x => x == "Corte"), "pas de Workflow Corte dans le pdf");
            Assert.AreEqual(2, mots.Count(x => x == "ETD"), "pas de ETD dans le pdf (1)");
            Assert.AreEqual(2 + 4, mots.Count(x => x == "23:00"), "pas de ETD dans le pdf (2)");
            // Trolley Label déjà compté

            //2ème tableau
            //-1ère ligne : galley location +aircraft + start date LPcart
            // Trolley Label déjà compté 3 fois
            Assert.AreEqual(2, mots.Count(x => x == aircraft), "pas de aircraft dans le pdf");
            Assert.AreEqual(2, mots.Count(x => x == from.ToString("d-MMM-yyyy", CultureInfo.GetCultureInfo("en-US"))), "pas de start date LPcart dans le pdf");
            //-2ème ligne : trolley name +galley code + LP cart comment
            //-comment : trolley comment
            // Trolley Label déjà compté 1 fois

            //3ème tableau
            //Front and Back(trolley scheme detail)
            Assert.AreEqual(4, mots.Count(x => x == "A00:" || x == "B00"), "pas de Front and Back dans le pdf (1)");
            Assert.AreEqual(4, mots.Count(x => x == "A10:" || x == "B10"), "pas de Front and Back dans le pdf (2)");
            Assert.AreEqual(4, mots.Count(x => x == "A01:" || x == "B01"), "pas de Front and Back dans le pdf (3)");
            Assert.AreEqual(4, mots.Count(x => x == "A11:" || x == "B11"), "pas de Front and Back dans le pdf (4)");
            //Trolley equipment -Trolley code
            // Trolley Label déjà compté 2 fois
            //LP cart name
            Assert.AreEqual(2, mots.Count(x => x == name), "pas de LP cart name dans le pdf");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PDF_servicesendguest()
        {
            //FIXME Cancelled non géré ici

            //Imprimer les "track sheets"
            var site = TestContext.Properties["SiteLpCart"].ToString();
            //Prepare
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string DocFileNamePdfBegin = "Print TrackSheet_-_";
            string DocFileNameZipBegin = "All_files_";

            //Auth
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.ClearDownloads();
            homePage.PurgeDownloads();

            //Act
            //Avoir des données disponibles.
            FlightPage flights = homePage.GoToFlights_FlightPage();
            flights.ResetFilter();
            flights.Filter(FilterType.Sites, site);
            Assert.IsTrue(flights.CheckTotalNumber() > 0);

            //1. cliquer sur "print track sheet"
            //2.Cliquer sur "print"
            PrintReportPage reportPage = flights.PrintReport(PrintType.TrackSheet, true);
            bool isPrintFlightLabels = reportPage.IsReportGenerated();
            reportPage.Close();
            reportPage.ClosePrintButton();
            Assert.IsTrue(isPrintFlightLabels, "Le rapport n'a pas été généré.");
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            reportPage.ClosePrintButton();
            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            // le tableau
            List<string> flightsPDF = new List<string>();
            List<string> flightsNonPDF = new List<string>();
            List<string> motsGuestServiceList = new List<string>();

            flights.ResetFilter();
            flights.Filter(FilterType.Sites, site);
            flights.PageSize("100");
            flights.ShowStatusMenu();
            flights.Filter(FilterType.Status, "None");
            // les services, les guest
            var flightsNo = WebDriver.FindElements(By.XPath("//*/tr[contains(@class,'row-dispatch-line')]/td[2]/input"));
            //65756339 - (984830)
            var flightsNo2 = WebDriver.FindElements(By.XPath("//*/tr[contains(@class,'row-dispatch-line')]/td[2]/div/input"));
            Console.WriteLine("None : " + flightsNo.Count + " ?= " + flights.CheckTotalNumber());
            foreach (var flight in flightsNo)
            {
                flightsNonPDF.Add(flight.GetAttribute("value"));
            }
            foreach (var flight in flightsNo2)
            {
                flightsNonPDF.Add(flight.GetAttribute("value"));
            }
            flights.ResetFilter();
            flights.Filter(FilterType.Sites, site);
            flights.PageSize("100");
            flights.ShowStatusMenu();
            flights.Filter(FilterType.Status, "Preval");
            flights.PageSize("100");
            // les services, les guest
            flightsNo = WebDriver.FindElements(By.XPath("//*/tr[contains(@class,'row-dispatch-line')]/td[2]/input"));
            //65756339 - (984830)
            flightsNo2 = WebDriver.FindElements(By.XPath("//*/tr[contains(@class,'row-dispatch-line')]/td[2]/div/input"));
            Console.WriteLine("Preval : " + flightsNo.Count + " ?= " + flights.CheckTotalNumber());
            List<string> flightsP = new List<string>();
            foreach (var flight in flightsNo)
            {
                flightsP.Add(flight.GetAttribute("value"));
            }
            foreach (var flight in flightsNo2)
            {
                flightsP.Add(flight.GetAttribute("value"));
            }
            foreach (var flight in flightsP)
            {
                //- vol avec guest-service statut dévalidé : pas d'infos dans le pdf : "No valid flights exist with thesefilters" 
                // if (flights.HasService(flight, new List<string>() { "Service for BOB" }))
                // P/V/I
                // statut P ou V ou I
                // Prévalidé ou validé ou invoiced
                if (flights.HasGuest(flight, motsGuestServiceList))
                {
                    // customer IsBob et customer.BobFlightsFillReconMode <> 1
                    flightsPDF.Add(flight);
                }
                else
                {
                    flightsNonPDF.Add(flight);
                }
            }
            flights.ResetFilter();
            flights.Filter(FilterType.Sites, site);
            flights.PageSize("100");
            flights.ShowStatusMenu();
            flights.Filter(FilterType.Status, "Valid");
            // les services, les guest
            flightsNo = WebDriver.FindElements(By.XPath("//*/tr[contains(@class,'row-dispatch-line')]/td[2]/input"));
            //65756339 - (984830)
            flightsNo2 = WebDriver.FindElements(By.XPath("//*/tr[contains(@class,'row-dispatch-line')]/td[2]/div/input"));
            Console.WriteLine("Valid : " + flightsNo.Count + " ?= " + flights.CheckTotalNumber());
            List<string> flightsPV = new List<string>();
            foreach (var flight in flightsNo)
            {
                flightsPV.Add(flight.GetAttribute("value"));
            }
            foreach (var flight in flightsNo2)
            {
                flightsPV.Add(flight.GetAttribute("value"));
            }
            foreach (var flight in flightsPV)
            {
                //- vol sans guest-service statut P/V : pas d'infos dans le pdf : "No valid flights exist with thesefilters" 
                if (flights.HasGuest(flight, motsGuestServiceList))
                {
                    flightsPDF.Add(flight);
                }
                else
                {
                    flightsNonPDF.Add(flight);
                }
            }
            flights.ResetFilter();
            flights.Filter(FilterType.Sites, site);
            flights.PageSize("100");
            flights.ShowStatusMenu();
            flights.Filter(FilterType.Status, "Invoice");
            // les services, les guest
            flightsNo = WebDriver.FindElements(By.XPath("//*/tr[contains(@class,'row-dispatch-line')]/td[2]/input"));
            //65756339 - (984830)
            flightsNo2 = WebDriver.FindElements(By.XPath("//*/tr[contains(@class,'row-dispatch-line')]/td[2]/div/input"));
            Console.WriteLine("Invoice : " + flightsNo.Count + " ?= " + flights.CheckTotalNumber());
            List<string> flightsPVI = new List<string>();
            foreach (var flight in flightsNo)
            {
                flightsPVI.Add(flight.GetAttribute("value"));
            }
            foreach (var flight in flightsNo2)
            {
                flightsPVI.Add(flight.GetAttribute("value"));
            }
            foreach (var flight in flightsPVI)
            {
                //- vol avec/sans guest-service statut P/V/I : infos dans le pdf
                flightsPDF.Add(flight);
            }

            flights.ResetFilter();
            flights.Filter(FilterType.Sites, site);
            flights.PageSize("100");

            var motsGuestServiceGroup = motsGuestServiceList.GroupBy(info => info)
                        .Select(group => new
                        {
                            Mot = group.Key,
                            Count = group.Count()
                        });


            //Le PDF imprimé devrait contenir:
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }

            foreach (var f in flightsPDF)
            {
                // 272500825_2024-11-12 => 272500825
                string f2 = f;
                if (f.Contains("_"))
                {
                    f2 = f.Substring(0, f.IndexOf("_"));
                }
                Assert.IsTrue(mots.Contains(f2), "flight pas dans PDF " + f2);
            }
            foreach (var f in flightsNonPDF)
            {
                // 272500825_2024-11-12 => 272500825
                string f2 = f;
                if (f.Contains("_"))
                {
                    f2 = f.Substring(0, f.IndexOf("_"));
                }
                Assert.IsFalse(mots.Contains(f2), "flight dans NonPDF " + f2);
            }

            //Vérifier le distingo des services entre les guest type sur le PDF
            //Vérifier que chaque guest type soient bien séparé sur le PDF

            foreach (var result in motsGuestServiceGroup)
            {
                if (result.Mot == "for") continue; //Observations: Lp for flight bob
                if (result.Mot == "") continue; // vide
                if (result.Mot == "TVS") continue; // 18 + 13 x (TVS)
                if (result.Mot == "BOB") continue; // 15 vs 11 : Service for BOB et guest BOB
                Console.WriteLine(result.Mot + " " + result.Count);
                Assert.AreEqual(mots.Count(x => x == result.Mot), result.Count, "mauvais nb " + result.Mot);
            }

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PDF_Greece()
        {
            //Prepare
            string valid = "Valid";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Delivery Note";
            string DocFileNameZipBegin = "All_files_";
            var deliveryNote = "Greece";
            string flightNumber = "";
            string site = TestContext.Properties["SiteACE"].ToString();
            // Arrange 
            HomePage homePage = LogInAsAdmin();

            //Act
            var paramPage = homePage.GoToParameters_GlobalSettings();
            var paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            var flightPage = new FlightPage(WebDriver, TestContext);
            flightPage.RemoveAllFromSelectedDeliveryNoteReport();
            #region refresh
            homePage.Navigate();
            paramPage = homePage.GoToParameters_GlobalSettings();
            paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            #endregion

            flightPage.SelectDeliveryNoteReport(deliveryNote);
            flightPage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.ShowStatusFilter();
            flightPage.Filter(FilterType.Sites, site);
            // il faut un avion validé avec un service (on le passe à P, on ajoute le service, on le repasse à V)
            if (flightPage.GetState(1) == "P")
            {
                // ajouter un guest+service
                FlightDetailsPage fDetails = flightPage.EditFirstFlight(flightPage.GetFirstFlightNumber());
                fDetails.AddGuestType("FC");
                fDetails.AddService("Service for flight");
                flightPage = fDetails.CloseViewDetails();
                // valider le premier flight
                flightPage.SetNewState("V");
            }

            flightPage.Filter(FilterType.Status, valid);
            if (flightPage.CheckTotalNumber() == 0)
            {
                flightPage.Filter(FilterType.Status, "Preval");
                flightPage.SetNewState("V");
                flightPage.Filter(FilterType.Status, valid);
                flightNumber = flightPage.GetFirstFlightNumber();
                flightPage.Filter(FilterType.SearchFlight, flightNumber);
            }
            else
            {
                flightNumber = flightPage.GetFirstFlightNumber();
                flightPage.Filter(FilterType.SearchFlight, flightNumber);
            }

            flightPage.ClearDownloads();
            flightPage.PurgeDownloads();
            var printReportPage = flightPage.PrintReport(PrintType.DeliveryNote, true);
            bool isPrintFlightLabels = printReportPage.IsReportGenerated();
            printReportPage.Close();
            Assert.IsTrue(isPrintFlightLabels, "Le rapport n'a pas été généré.");

            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);

            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            //Le PDF imprimé devrait contenir:
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            string motsPlat = "";
            foreach (Page p in document.GetPages())
            {
                motsPlat = motsPlat + p.Text.Replace(" ", "");
            }
            //Assert
            bool isFlightExist = motsPlat.Contains(flightNumber);
            Assert.IsTrue(isFlightExist, "Les données du Pdf ne correspondent pas à la grille.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PDF_New()
        {
            //Prepare
            string valid = "Valid";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Delivery Note";
            string DocFileNameZipBegin = "All_files_";
            var deliveryNote = "New";
            string site = TestContext.Properties["SiteACE"].ToString();

            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var paramPage = homePage.GoToParameters_GlobalSettings();
            var paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();

            var flightPage = new FlightPage(WebDriver, TestContext);
            flightPage.RemoveAllFromSelectedDeliveryNoteReport();
            #region refresh
            homePage.Navigate();
            paramPage = homePage.GoToParameters_GlobalSettings();
            paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            #endregion

            flightPage.SelectDeliveryNoteReport(deliveryNote);
            flightPage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.ShowStatusFilter();
            flightPage.Filter(FilterType.Sites, site);
            // il faut un avion validé avec un service (on le passe à P, on ajoute le service, on le repasse à V)
            if (flightPage.GetState(1) == "P")
            {
                // ajouter un guest+service
                FlightDetailsPage fDetails = flightPage.EditFirstFlight(flightPage.GetFirstFlightNumber());
                fDetails.AddGuestType("FC");
                fDetails.AddService("Service for flight");
                flightPage = fDetails.CloseViewDetails();
                // valider le premier flight
                flightPage.SetNewState("V");
            }
            flightPage.Filter(FilterType.Status, valid);
            var flightNumber = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNumber);
            flightPage.ClearDownloads();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            //clean directory
            foreach (var file in taskDirectory.GetFiles())
            {
                file.Delete();
            }
            var printReportPage = flightPage.PrintReport(PrintType.DeliveryNote, true);
            bool isPrintFlightLabels = printReportPage.IsReportGenerated();
            printReportPage.Close();
            //Assert
            Assert.IsTrue(isPrintFlightLabels, "Le rapport n'a pas été généré.");

            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);

            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            //Le PDF imprimé devrait contenir:
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            string mots = "";
            foreach (Page p in document.GetPages())
            {
                mots += p.Text.Replace(" ", "");
            }
            //Assert
            bool isFlightExist = mots.Contains(flightNumber);
            Assert.IsTrue(isFlightExist, "Les données du Pdf ne correspondent pas à la grille.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PDF_New2()
        {
            //Prepare
            string valid = "Valid";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Delivery Note";
            string DocFileNameZipBegin = "All_files_";
            var deliveryNote = "New2";
            string site = TestContext.Properties["SiteACE"].ToString();

            string serviceName = "Service for flight";
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            DateTime fromDate = DateUtils.Now.AddMonths(-30);
            DateTime toDate = DateUtils.Now.AddMonths(+30);
            string datasheet = "Renaud Datasheet";
            string customer = "TVS - SMARTWINGS, A.S. (TVS)"; //"SMARTWINGS, A.S. (TVS)";
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string flightNumber = new Random().Next().ToString();
            //arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var paramPage = homePage.GoToParameters_GlobalSettings();
            var paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();

            var flightPage = new FlightPage(WebDriver, TestContext);
            flightPage.RemoveAllFromSelectedDeliveryNoteReport();
            #region refresh
            homePage.Navigate();
            paramPage = homePage.GoToParameters_GlobalSettings();
            paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            #endregion

            flightPage.SelectDeliveryNoteReport(deliveryNote);
            flightPage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.ShowStatusFilter();
            //flightPage.Filter(FilterType.Sites, site);
            flightPage.GoToFlights_FlightPage();
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
            flightPage.ResetFilter();
            flightPage.ShowStatusFilter();
            flightPage.Filter(FilterType.Sites, site);
            flightPage.Filter(FilterType.SearchFlight, flightNumber);
            // il faut un avion validé avec un service (on le passe à P, on ajoute le service, on le repasse à V)
            if (flightPage.GetState(1) == "P")
            {
                // ajouter un guest+service
                FlightDetailsPage fDetails = flightPage.EditFirstFlight(flightPage.GetFirstFlightNumber());
                fDetails.AddGuestType("FC");
                fDetails.AddService("Service for flight");
                flightPage = fDetails.CloseViewDetails();
                // valider le premier flight
                flightPage.SetNewState("V");
            }
            flightPage.Filter(FilterType.Status, valid);
            //var flightNumber = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNumber);

            flightPage.ClearDownloads();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            //clean directory
            foreach (var file in taskDirectory.GetFiles())
            {
                file.Delete();
            }
            var printReportPage = flightPage.PrintReport(PrintType.DeliveryNote, true);
            bool isPrintFlightLabels = printReportPage.IsReportGenerated();
            printReportPage.Close();
            Assert.IsTrue(isPrintFlightLabels, "Le rapport n'a pas été généré.");
            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            //Le PDF imprimé devrait contenir:
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            string mots = "";
            foreach (Page p in document.GetPages())
            {
                mots += p.Text.Replace(" ", "");
            }
            //Assert
            bool isFlightExist = mots.Contains(flightNumber);
            Assert.IsTrue(isFlightExist, "Les données du Pdf ne correspondent pas à la grille.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PDF_Valorized()
        {
            //Prepare
            string valid = "Valid";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Delivery Note";
            string DocFileNameZipBegin = "All_files_";
            var deliveryNote = "Valorized";
            string site = TestContext.Properties["SiteACE"].ToString();

            //string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            string serviceName = "Service for flight";
            //Prepare
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = "TVS - SMARTWINGS, A.S. (TVS)"; //"SMARTWINGS, A.S. (TVS)";

            string flightNumber = new Random().Next().ToString();

            //string serviceName = "servForFlight" + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            DateTime fromDate = DateUtils.Now.AddMonths(-30);
            DateTime toDate = DateUtils.Now.AddMonths(+30);
            string datasheet = "Renaud Datasheet";

            //arrange
            HomePage homePage = LogInAsAdmin();
         
            //Act
            var paramPage = homePage.GoToParameters_GlobalSettings();
            var paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();

            var flightPage = new FlightPage(WebDriver, TestContext);
            flightPage.RemoveAllFromSelectedDeliveryNoteReport();
            #region refresh
            homePage.Navigate();
            paramPage = homePage.GoToParameters_GlobalSettings();
            paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            #endregion

            flightPage.SelectDeliveryNoteReport(deliveryNote);

            var servicesPage = homePage.GoToCustomers_ServicePage();
            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);
            ServicePricePage pricePage = servicesPage.ClickOnFirstService();
            pricePage.UnfoldAll();
            var priceModalPage = pricePage.EditFirstPrice(site, "");
            pricePage = priceModalPage.FillFields_EditCustomerPrice(site, customer, fromDate, toDate);

            WebDriver.Navigate().Refresh();
            pricePage.UnfoldAll();
            pricePage.SetFirstPriceDataSheet(datasheet);
            flightPage.GoToFlights_FlightPage();
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
            flightPage.ResetFilter();
            flightPage.ShowStatusFilter();
            flightPage.Filter(FilterType.Sites, site);
            flightPage.Filter(FilterType.SearchFlight, flightNumber);
            // il faut un avion validé avec un service (on le passe à P, on ajoute le service, on le repasse à V)
            if (flightPage.GetState(1) == "P")
            {
                // ajouter un guest+service
                FlightDetailsPage fDetails = flightPage.EditFirstFlight(flightPage.GetFirstFlightNumber());
                fDetails.AddGuestType("FC");
                fDetails.AddService(serviceName);
                flightPage = fDetails.CloseViewDetails();
                // valider le premier flight
                flightPage.SetNewState("V");
            }
            flightPage.Filter(FilterType.Status, valid);
            //var flightNumber = flightPage.GetFirstFlightNumber();

            flightPage.ClearDownloads();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            //clean directory
            foreach (var file in taskDirectory.GetFiles())
            {
                file.Delete();
            }
            var printReportPage = flightPage.PrintReport(PrintType.DeliveryNote, true, true);
            bool isPrintFlightLabels = printReportPage.IsReportGenerated();
            printReportPage.Close();
            //Assert
            Assert.IsTrue(isPrintFlightLabels, "Le rapport n'a pas été généré.");

            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);

            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            //Assert
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            //Le PDF imprimé devrait contenir:
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            string mots = "";
            foreach (Page p in document.GetPages())
            {
                mots += p.Text.Replace(" ", "");
            }
            //Assert
            bool isFlightExist = mots.Contains(flightNumber);
            Assert.IsTrue(isFlightExist, "Les données du Pdf ne correspondent pas à la grille.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PDF_Spain()
        {
            //Prepare
            string valid = "Valid";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Delivery Note";
            string DocFileNameZipBegin = "All_files_";
            var deliveryNote = "Spain";
            string site = TestContext.Properties["SiteACE"].ToString();

            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var paramPage = homePage.GoToParameters_GlobalSettings();
            var paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();

            var flightPage = new FlightPage(WebDriver, TestContext);
            flightPage.RemoveAllFromSelectedDeliveryNoteReport();
            #region refresh
            homePage.Navigate();
            paramPage = homePage.GoToParameters_GlobalSettings();
            paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            #endregion
            flightPage.SelectDeliveryNoteReport(deliveryNote);
            flightPage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.ShowStatusFilter();
            flightPage.Filter(FilterType.Sites, site);
            if (flightPage.GetState(1) == "P")
            {
                // ajouter un guest+service
                FlightDetailsPage fDetails = flightPage.EditFirstFlight(flightPage.GetFirstFlightNumber());
                fDetails.AddGuestType("FC");
                fDetails.AddService("Service for flight");
                flightPage = fDetails.CloseViewDetails();
                // valider le premier flight
                flightPage.SetNewState("V");
            }
            flightPage.Filter(FilterType.Status, valid);
            var flightNumber = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNumber);

            flightPage.ClearDownloads();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            //clean directory
            foreach (var file in taskDirectory.GetFiles())
            {
                file.Delete();
            }
            var printReportPage = flightPage.PrintReport(PrintType.DeliveryNote, true);
            bool isPrintFlightLabels = printReportPage.IsReportGenerated();
            printReportPage.Close();
            //Assert
            Assert.IsTrue(isPrintFlightLabels, "Le rapport n'a pas été généré.");
            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);

            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            //Assert
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            //Le PDF imprimé devrait contenir:
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            string mots = "";
            foreach (Page p in document.GetPages())
            {
                mots += p.Text.Replace(" ", "");
            }
            //Assert
            bool isFlightExist = mots.Contains(flightNumber);
            Assert.IsTrue(isFlightExist, "Les données du Pdf ne correspondent pas à la grille.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PDF_Old()
        {
            //Prepare
            string valid = "Valid";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Delivery Note";
            string DocFileNameZipBegin = "All_files_";
            var deliveryNote = "Old";
            string site = TestContext.Properties["SiteACE"].ToString();
            //string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            string serviceName = "Service for flight";
            //Prepare
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = "TVS - SMARTWINGS, A.S. (TVS)"; //"SMARTWINGS, A.S. (TVS)";

            string flightNumber = new Random().Next().ToString();

            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var paramPage = homePage.GoToParameters_GlobalSettings();
            var paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();

            var flightPage = new FlightPage(WebDriver, TestContext);
            flightPage.RemoveAllFromSelectedDeliveryNoteReport();
            #region refresh
            homePage.Navigate();
            paramPage = homePage.GoToParameters_GlobalSettings();
            paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            #endregion
            flightPage.SelectDeliveryNoteReport(deliveryNote);
            flightPage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            flightPage.ShowStatusFilter();
            flightPage.Filter(FilterType.Sites, site);
            
            flightPage.Filter(FilterType.SearchFlight, flightNumber);

            // il faut un avion validé avec un service (on le passe à P, on ajoute le service, on le repasse à V)
            if (flightPage.GetState(1) == "P")
            {
                // ajouter un guest+service
                FlightDetailsPage fDetails = flightPage.EditFirstFlight(flightPage.GetFirstFlightNumber());
                fDetails.AddGuestType("FC");
                //fDetails.AddService("Service for flight");
                fDetails.AddService(serviceName);
                flightPage = fDetails.CloseViewDetails();
                // valider le premier flight
                flightPage.SetNewState("V");
            }
            flightPage.Filter(FilterType.Status, valid);
            //var flightNumber = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNumber);

            flightPage.ClearDownloads();
            flightPage.PurgeDownloads();
            var printReportPage = flightPage.PrintReport(PrintType.DeliveryNote, true);
            bool isPrintFlightLabels = printReportPage.IsReportGenerated();
            printReportPage.Close();
            //Assert
            Assert.IsTrue(isPrintFlightLabels, "Le rapport n'a pas été généré.");
            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);

            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            //Assert
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");
            //Le PDF imprimé devrait contenir:
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }
            //Assert
            bool isFlightExist = mots.Contains(flightNumber);
            Assert.IsTrue(isFlightExist, "Les données du Pdf ne correspondent pas à la grille.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PDF_Peru()
        {
            //Prepare
            string valid = "Valid";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Delivery Note";
            string DocFileNameZipBegin = "All_files_";
            var deliveryNote = "Peru";
            string site = TestContext.Properties["SiteACE"].ToString();
           
            //arrange
            HomePage homePage = LogInAsAdmin();
           
            //Act
            var paramPage = homePage.GoToParameters_GlobalSettings();
            var paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            var flightPage = new FlightPage(WebDriver, TestContext);
            flightPage.RemoveAllFromSelectedDeliveryNoteReport();
            #region refresh
            homePage.Navigate();
            paramPage = homePage.GoToParameters_GlobalSettings();
            paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            #endregion
            flightPage.SelectDeliveryNoteReport(deliveryNote);
            flightPage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.ShowStatusFilter();
            flightPage.Filter(FilterType.Sites, "ACE");
            // il faut un avion validé avec un service (on le passe à P, on ajoute le service, on le repasse à V)
            if (flightPage.GetState(1) == "P")
            {
                // ajouter un guest+service
                FlightDetailsPage fDetails = flightPage.EditFirstFlight(flightPage.GetFirstFlightNumber());
                fDetails.AddGuestType("FC");
                fDetails.AddService("Service for flight");
                flightPage = fDetails.CloseViewDetails();
                // valider le premier flight
                flightPage.SetNewState("V");
            }
            flightPage.Filter(FilterType.Status, valid);
            var flightNumber = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNumber);

            flightPage.ClearDownloads();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            //clean directory
            foreach (var file in taskDirectory.GetFiles())
            {
                file.Delete();
            }
            var printReportPage = flightPage.PrintReport(PrintType.DeliveryNote, true);
            bool isPrintFlightLabels = printReportPage.IsReportGenerated();
            printReportPage.Close();
            //Assert
            Assert.IsTrue(isPrintFlightLabels, "Le rapport n'a pas été généré.");
            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            //Le PDF imprimé devrait contenir:
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            string mots = "";
            foreach (Page p in document.GetPages())
            {
                mots += p.Text.Replace(" ", "");
            }
            //Assert
            bool isFlightExist = mots.Contains(flightNumber);
            Assert.IsTrue(isFlightExist, "Les données du Pdf ne correspondent pas à la grille.");
        }



        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_MassiveDelete()
        {
            //Prepare
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = "CAT Genérico"; //"SMARTWINGS, A.S. (TVS)";

            string flightNumber = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FilterType.Sites, siteFrom);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            flightPage.ResetFilter();

            flightPage.MassiveDeleteMenus(flightNumber, siteFrom, customer);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);

            //Assert
            Assert.AreEqual(0, flightPage.CheckTotalNumber(), "La massive delete ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_DeleteMultipleFlights()
        {
            //Prepare
            string siteFrom = TestContext.Properties["CustomerCodeSchedule"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string flightNumber = new Random().Next().ToString();
            string customer = "SMARTWINGS, A.S. (TVS)";

            string flightNumber2 = flightNumber + 1;


            DateTime DateFrom = DateUtils.Now;
            DateTime DateTo = DateUtils.Now.AddDays(1);
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FilterType.Sites, siteFrom);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
            flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, siteFrom, siteTo);
            flightPage.ResetFilter();

            flightPage.MultipleDelete(customer, siteFrom, DateFrom, DateTo);
            flightPage.Filter(FilterType.Sites, siteFrom);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            //Assert
            Assert.AreEqual(0, flightPage.CheckTotalNumber(), "La multiple delete ne fonctionne pas.");
            flightPage.Filter(FilterType.Sites, siteFrom);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber2);
            //Assert
            Assert.AreEqual(0, flightPage.CheckTotalNumber(), "La multiple delete ne fonctionne pas.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_DetailFlightDeleteService()
        {
            //Prepare
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string flightNumber = new Random().Next().ToString();
            string customer = "SMARTWINGS, A.S. (TVS)";
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            string serviceName2 = TestContext.Properties["ServiceBob"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                //Edit the first flight
                var editPage = flightPage.EditFirstFlight(flightNumber);
                editPage.AddGuestType();
                editPage.AddService(serviceName);
                editPage.AddService(serviceName2);
                var createdService = editPage.CountServices();
                Assert.AreEqual(2, createdService, MessageErreur.SERVICE_NON_AJOUTE);

                flightPage = editPage.CloseViewDetails();
                flightPage.EditFirstFlight(flightNumber);
                editPage.DeleteService();
                editPage.CloseModal();

                flightPage.Filter(FilterType.SearchFlight, flightNumber);
                flightPage.EditFirstFlight(flightNumber);
                //Assert
                var finalResult = editPage.CountServices();
                Assert.AreEqual(1, finalResult, "Le service n'a pas été supprimé.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_VueCalendarUnfoldAll()
        {

            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();

            //search flight
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.SearchFlight, flightNumber);
            flightPage.WaitPageLoading();

            //edit flight
            var editPage = flightPage.EditFirstFlight(flightNumber);

            //add guest type
            editPage.AddGuestType("BOB");

            //add service
            editPage.AddService(serviceName);
            editPage.SetFinalQty("0");
            editPage.WaitPageLoading();
            editPage.CloseModal();

            //Get Calendar View
            flightPage.GetCalendarView();
            flightPage.UnfoldAll();
            var isUnfoldAll = flightPage.IsUnfoldAll();
            var isDetail = flightPage.IsDetail();

            //Assert
            Assert.IsTrue(isUnfoldAll, "Le détail des flights n'est pas affiché.");
            Assert.IsTrue(isDetail, "Le détail des flights n'est pas affiché.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_FilterEquipementType_Atlas()
        {
            //prepare
            string aircraftName = "Aircarft"+ new Random().Next();
            string equipementType = "A";
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string flightNumber = new Random().Next().ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var parametersFlightPage = homePage.GoToParametres_Flights();
            var aircraftPage =  parametersFlightPage.GoToAircraft();
            try
            {
                // cree un aircraft avec euipement ATLAS
                aircraftPage.CreateAircraftModelPage();
                aircraftPage.FillFieldsAircraftModelPage(aircraftName,equipementType);
                aircraftPage.Save();
                // crée un flight
                var flightPage = homePage.GoToFlights_FlightPage();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraftName, siteFrom, siteTo);
                // apply the filter
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNumber);
                flightPage.WaitPageLoading();
                flightPage.ShowEquipmentTypeMenu();
                flightPage.Filter(FilterType.EQUIPMENTTYPE_Atlas, false);
                flightPage.WaitPageLoading();
                var numberOfFlight = flightPage.GetCountResult();
                Assert.AreEqual(0, numberOfFlight, "le filtre sur euipement type atlas ne s'apllique pas");
                flightPage.Filter(FilterType.EQUIPMENTTYPE_Atlas, true);
                flightPage.WaitPageLoading();
                numberOfFlight = flightPage.GetCountResult();
                Assert.AreEqual(1, numberOfFlight, "le filtre sur euipement type atlas ne s'apllique pas");

            }
            finally
            {
                // delete flight
                var flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
                // delete aircraft
                parametersFlightPage = homePage.GoToParametres_Flights();
                aircraftPage = parametersFlightPage.GoToAircraft();
                aircraftPage.FilterBySearch(aircraftName);
                aircraftPage.WaitPageLoading();
                aircraftPage.DeleteFirstAircraft();
            }            
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_FilterEquipementType_KSSU()
        {
            //prepare
            string aircraftName = "Aircarft" + new Random().Next();
            string equipementType = "K";
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string flightNumber = new Random().Next().ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var parametersFlightPage = homePage.GoToParametres_Flights();
            var aircraftPage = parametersFlightPage.GoToAircraft();
            try
            {
                // cree un aircraft avec euipement KSSU
                aircraftPage.CreateAircraftModelPage();
                aircraftPage.FillFieldsAircraftModelPage(aircraftName, equipementType);
                aircraftPage.Save();
                // crée un flight
                var flightPage = homePage.GoToFlights_FlightPage();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraftName, siteFrom, siteTo);
                // apply the filter
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNumber);
                flightPage.WaitPageLoading();
                flightPage.ShowEquipmentTypeMenu();
                flightPage.Filter(FilterType.EQUIPMENTTYPE_KSSU, false);
                flightPage.WaitPageLoading();
                var numberOfFlight = flightPage.GetCountResult();
                Assert.AreEqual(0, numberOfFlight, "le filtre sur euipement type Kssu ne s'apllique pas");
                flightPage.Filter(FilterType.EQUIPMENTTYPE_KSSU, true);
                flightPage.WaitPageLoading();
                numberOfFlight = flightPage.GetCountResult();
                Assert.AreEqual(1, numberOfFlight, "le filtre sur euipement type Kssu ne s'apllique pas");

            }
            finally
            {
                // delete flight
                var flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
                // delete aircraft
                parametersFlightPage = homePage.GoToParametres_Flights();
                aircraftPage = parametersFlightPage.GoToAircraft();
                aircraftPage.FilterBySearch(aircraftName);
                aircraftPage.WaitPageLoading();
                aircraftPage.DeleteFirstAircraft();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_FilterEquipementType_Unknown()
        {
            //prepare
            string aircraftName = "Aircarft" + new Random().Next();
            string equipementType = "UNKNOWN";
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string flightNumber = new Random().Next().ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var parametersFlightPage = homePage.GoToParametres_Flights();
            var aircraftPage = parametersFlightPage.GoToAircraft();
            try
            {
                // cree un aircraft avec euipement UNKNOWN
                aircraftPage.CreateAircraftModelPage();
                aircraftPage.FillFieldsAircraftModelPage(aircraftName, equipementType);
                aircraftPage.Save();
                // crée un flight
                var flightPage = homePage.GoToFlights_FlightPage();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraftName, siteFrom, siteTo);
                // apply the filter
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNumber);
                flightPage.WaitPageLoading();
                flightPage.ShowEquipmentTypeMenu();
                flightPage.Filter(FilterType.EQUIPMENTTYPE_Unknown, false);
                flightPage.WaitPageLoading();
                var numberOfFlight = flightPage.GetCountResult();
                Assert.AreEqual(0, numberOfFlight, "le filtre sur euipement type Unknown ne s'apllique pas");
                flightPage.Filter(FilterType.EQUIPMENTTYPE_Unknown, true);
                flightPage.WaitPageLoading();
                numberOfFlight = flightPage.GetCountResult();
                Assert.AreEqual(1, numberOfFlight, "le filtre sur euipement type Unknown ne s'apllique pas");

            }
            finally
            {
                // delete flight
                var flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
                // delete aircraft
                parametersFlightPage = homePage.GoToParametres_Flights();
                aircraftPage = parametersFlightPage.GoToAircraft();
                aircraftPage.FilterBySearch(aircraftName);
                aircraftPage.WaitPageLoading();
                aircraftPage.DeleteFirstAircraft();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_FilterActivePrice()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.SetDateState(new DateTime(2024, 4, 24));
            var numberbeforeFilter = flightPage.CheckTotalNumber();
            flightPage.WaiForLoad();
            flightPage.Filter(FlightPage.FilterType.noActivePrice, true);
            flightPage.WaitPageLoading();
            var numberafterFilter = flightPage.CheckTotalNumber();

            //Assert 
            // une vol associé à ce service expiré existant
            Assert.AreNotEqual(numberbeforeFilter, numberafterFilter);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_FilterShowNotCounted()
        {
            //Prepare
            string flightBobCust = "Darabata";
            string site = "MAD";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.SetDateState(new DateTime(2024, 4, 21));

            //Assert
            // filght bob avec propriété AllRelatedTrolleysAreCounted vrai
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightBobCust);
            var numberFlightFoundCounted = flightPage.CheckTotalNumber();
            Assert.AreEqual(1, numberFlightFoundCounted);

            //filtré par avec filtre not counted 
            flightPage.Filter(FlightPage.FilterType.showNotCounted, true);
            Thread.Sleep(4000);
            var numberFlightFoundbitCoounted = flightPage.CheckTotalNumber();

            Assert.AreEqual(0, numberFlightFoundbitCoounted);

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_FilterShowBobOfOnly()
        {
            //Prepare
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string noCustomerBob = TestContext.Properties["CustomerLPFlight"].ToString();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customerBob);
            var customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            var icao = customerGeneralInformationsPage.GetIcao();

            if (!customerGeneralInformationsPage.IsBoBVisible())
            {
                customerGeneralInformationsPage.ActiveBuyOnBoard();
            }

            //Act
            var flightList = homePage.GoToFlights_FlightPage();
            flightList.ResetFilter();
            flightList.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightModal = flightList.FlightCreatePage();
            flightModal.FillField_CreatNewFlight(flightNumber, noCustomerBob, aircraft, siteFrom, siteTo);


            flightList.ResetFilter();
            flightList.SetBobTop();
            flightList.Filter(FlightPage.FilterType.CustomerBob, customer);

            if (!flightList.PageSizeEqualsTo100())
            {
                flightList.PageSize("8");
                flightList.PageSize("100");
            }

            //Assert
            Assert.IsTrue(flightList.VerifyCustomer(icao), MessageErreur.FILTRE_ERRONE, "CustomersBob");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_sitesearch()
        {
            //prepare
            DateTime dateTest = DateTime.Now;


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            var modalDelete = flightPage.ClickMassiveDelete();

            modalDelete.SetDateFromAndTo(dateTest, dateTest);
            //select site to filter 
            modalDelete.SelectSiteToFilter();
            modalDelete.ClickSearchButton();
            var result = modalDelete.CheckIfAllLineContainTheSite("ACE");

            //Assert
            Assert.IsTrue(result, "Anomalie, presence site autre que site selectionné");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_searchByDate()
        {
            //prepare
            DateTime dateTestDeb = DateTime.Today;
            DateTime dateTestFin = dateTestDeb.AddDays(2);


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            var modalDelete = flightPage.ClickMassiveDelete();
            modalDelete.SetDateFromAndTo(dateTestDeb, dateTestFin);
            modalDelete.ClickSearchButton();
            var listeDateReuslt = modalDelete.GetCurrentDatesResult().ToList();
            var result = modalDelete.CheckIfDateIsGood(dateTestDeb, dateTestFin, listeDateReuslt);

            //Assert
            Assert.IsTrue(result, "Anomalie, presence date no compris dans filtre");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_inactivecustomercheck()
        {
            string inactiveCustomersName = "Inactive -";
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();

            massiveDeleteModal.ShowInactiveCustomers();

            massiveDeleteModal.SelectCustomer(inactiveCustomersName, true);
            massiveDeleteModal.ClickSearchButton();
            massiveDeleteModal.PageSize("50");
            var customers = massiveDeleteModal.CustomersWithoutDeplicates();
            massiveDeleteModal.CloseMassiveDeleteModal();
            homePage.Navigate();
            CustomerPage customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.ResetFilters();
            customerPage.Filter(CustomerPage.FilterType.ShowInactive, true);
            foreach (var customer in customers)
            {
                customerPage.Filter(CustomerPage.FilterType.Search, customer);
                var inactiveCustomer = customerPage.CheckTotalNumber();
                Assert.IsTrue(inactiveCustomer > 0, "Le show inactive customer ne fonctionne pas.");
            }

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_customersearch()
        {
            string customerName = "2 EXCEL AVIATION LTD";
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();

            massiveDeleteModal.SelectCustomer(customerName);

            massiveDeleteModal.ClickSearchButton();

            Assert.IsTrue(massiveDeleteModal.VerifyCustomerSearch(customerName), "Le search customer ne fonctionne pas.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_inactivecustomersearch()
        {
            string inactiveCustomerName = "Inactive - CEIBA INTERCONTINENTAL";
            string customerName = "CEIBA INTERCONTINENTAL";
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();

            massiveDeleteModal.ShowInactiveCustomers();

            massiveDeleteModal.SelectCustomer(inactiveCustomerName);

            massiveDeleteModal.ClickSearchButton();

            Assert.IsTrue(massiveDeleteModal.VerifyCustomerSearch(customerName), "Le search inactive customer ne fonctionne pas.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_organizebysite()
        {
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();

            massiveDeleteModal.ClickSearchButton();

            massiveDeleteModal.ClickOnSiteHeader();
            var firstComparison = massiveDeleteModal.VerifySiteCodeSort(SortType.Ascending);

            massiveDeleteModal.ClickOnSiteHeader();
            var secondComparison = massiveDeleteModal.VerifySiteCodeSort(SortType.Descending);

            Assert.IsTrue(firstComparison, " Les Sites ne sont pas ordonnés ( de A à Z )");
            Assert.IsTrue(secondComparison, " Les Sites ne sont pas ordonnés ( de Z à A )");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_organizebycustomer()
        {
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();

            massiveDeleteModal.ClickSearchButton();

            massiveDeleteModal.ClickOnCustomerHeader();
            var firstComparison = massiveDeleteModal.VerifyCustomerCodeSort(SortType.Ascending);

            massiveDeleteModal.ClickOnCustomerHeader();
            var secondComparison = massiveDeleteModal.VerifyCustomerCodeSort(SortType.Descending);

            Assert.IsTrue(firstComparison, " Les Customers ne sont pas ordonnés ( de A à Z )");
            Assert.IsTrue(secondComparison, " Les Customers ne sont pas ordonnés ( de Z à A )");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_inactivesitescheck()
        {
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();

            massiveDeleteModal.ShowInactiveSites();

            massiveDeleteModal.SelectSite("Inactive -");

            massiveDeleteModal.ClickOnSitesFilter();

            Assert.IsTrue(massiveDeleteModal.IsInactive(), "Le show inactive customer ne fonctionne pas.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_inactivesitessearch()
        {

            string site = "MAD";
            string inactive_site = "Inactive - MAD";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);

            homePage.Navigate();
            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.DeactivateFirstSiteInList();

            homePage.Navigate();
            ParametersUser userParameter = homePage.GoToParameters_User();
            userParameter.SearchAndSelectUser("wr.testauto");
            userParameter.ClickOnAffectedSite();
            userParameter.ActivateSite(site);
            try
            {
                homePage.Navigate();
                var flightPage = homePage.GoToFlights_FlightPage();
                FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();

                massiveDeleteModal.ShowInactiveSites();

                massiveDeleteModal.SelectSite(inactive_site);
                massiveDeleteModal.ClickSearchButton();

                int totalpage = 0;
                var isData = massiveDeleteModal.CheckData();
                if (isData)
                {
                    massiveDeleteModal.PageSize("100");
                    totalpage = massiveDeleteModal.CheckTotalPageNumber();
                }

                if (totalpage > 0)
                {
                    Assert.IsTrue(massiveDeleteModal.VerifySiteSearch(site), "erreur de filtrage par site");
                    massiveDeleteModal.GoToPage(totalpage);
                    Assert.IsTrue(massiveDeleteModal.VerifySiteSearch(site), "erreur de filtrage par site");
                }
            }
            catch
            {
                homePage.Navigate();
                siteParameterPage = homePage.GoToParameters_Sites();
                siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
                siteParameterPage.ActivateFirstSiteInList();

                homePage.Navigate();
                userParameter = homePage.GoToParameters_User();
                userParameter.SearchAndSelectUser("wr.testauto");
                userParameter.ClickOnAffectedSite();
                userParameter.ActivateSite(site);
            }
            homePage.Navigate();
            siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.ActivateFirstSiteInList();

            homePage.Navigate();
            userParameter = homePage.GoToParameters_User();
            userParameter.SearchAndSelectUser("wr.testauto");
            userParameter.ClickOnAffectedSite();
            userParameter.ActivateSite(site);
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_organizebyflight()
        {
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();

            massiveDeleteModal.ClickSearchButton();

            massiveDeleteModal.ClickOnFlightHeader();
            var firstComparison = massiveDeleteModal.VerifyFlightCodeSort(SortType.Ascending);

            massiveDeleteModal.ClickOnFlightHeader();
            var secondComparison = massiveDeleteModal.VerifyFlightCodeSort(SortType.Descending);

            Assert.IsTrue(firstComparison, " Les Flights ne sont pas ordonnés ( de A à Z )");
            Assert.IsTrue(secondComparison, " Les Flights ne sont pas ordonnés ( de Z à A )");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_OrganizeByBarset()
        {
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();

            massiveDeleteModal.ClickSearchButton();

            massiveDeleteModal.ClickOnBarsetHeader();
            var firstComparison = massiveDeleteModal.VerifyOrganizeByBarset("No");

            massiveDeleteModal.ClickOnBarsetHeader();
            var secondComparison = massiveDeleteModal.VerifyOrganizeByBarset("Yes");

            Assert.IsTrue(firstComparison && secondComparison, " Les Barsets ne sont pas organisés");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_pagination()
        {
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();
            massiveDeleteModal.ClickSearchButton();
            massiveDeleteModal.ClickSelectAllButton();

            int tot = massiveDeleteModal.CheckTotalSelectCount();

            massiveDeleteModal.PageSize("16");
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("16"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 16 ? 16 : tot, "Pagination ne fonctionne pas.");

            massiveDeleteModal.PageSize("30");
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("30"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 30 ? 30 : tot, "Pagination ne fonctionne pas.");

            massiveDeleteModal.PageSize("50");
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("50"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 50 ? 50 : tot, "Pagination ne fonctionne pas.");

            massiveDeleteModal.PageSize("100");
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("100"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 100 ? 100 : tot, "Pagination ne fonctionne pas.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_namesearch()
        {
            string flightName = "1832169303";
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();

            massiveDeleteModal.SetFlightName(flightName);

            massiveDeleteModal.ClickSearchButton();

            Assert.IsTrue(massiveDeleteModal.VerifyServiceNameSearch(flightName), "La recherche par flight name ne fonctionne pas.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_OrganizeByDate()
        {
            string flightName = "17020";
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();

            massiveDeleteModal.SetFlightName(flightName);

            massiveDeleteModal.ClickSearchButton();
            var sort = massiveDeleteModal.GetSortType();
            massiveDeleteModal.ClickOnDateHeader();
            var sortAfterClick = massiveDeleteModal.GetSortType();

            Assert.IsTrue(sort != sortAfterClick, "Le sort by date flight ne fonctionne pas.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_link()
        {
            Random rand = new Random();
            DateTime dateFin = DateTime.Today;
            DateTime dateDebut = dateFin.AddYears(-3);

            int range = (DateTime.Today - dateDebut.AddYears(1)).Days;
            dateDebut = dateDebut.AddDays(rand.Next(range));

            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();
            massiveDeleteModal.SetDateFromAndTo(dateDebut, dateFin);
            massiveDeleteModal.ClickSearchButton();

            string date = massiveDeleteModal.GetFirstRowDate();
            string site = massiveDeleteModal.GetFirstRowSite();
            Assert.IsTrue(date != null && site != null, "Aucun résultat de Flight trouvé pour une recherche de massive delete.");

            bool linkIsClicked = massiveDeleteModal.ClickOnResultLinkAndCheckFlight(date, site);
            Assert.IsTrue(linkIsClicked, "Lien du résultat vers le flight non opérationnel.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedeletes()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
            massiveDeleteFlightPage.ClickSearchButton();
            var firstFlightNumber = massiveDeleteFlightPage.GetFirstFlightNumber();
            massiveDeleteFlightPage.SelectFirstFlight();
            massiveDeleteFlightPage.Delete();
            flightPage.Filter(FilterType.SearchFlight, firstFlightNumber);
            var result = flightPage.CheckTotalNumber();
            Assert.IsTrue(result == 0, "flight not deleted");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_selectall()
        {
            //prepare
            DateTime dateTest = DateTime.Now.AddDays(-4);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            var modalDelete = flightPage.ClickMassiveDelete();

            modalDelete.SetDateFromAndTo(dateTest, dateTest);
            //select site to filter 
            modalDelete.SelectSiteToFilter();
            modalDelete.ClickSearchButton();
            modalDelete.PageSize("100");
            // select all result 
            modalDelete.ClickSelectAllButton();
            var allChecked = modalDelete.IsSelectedEachFlight();
            //Assert
            Assert.IsTrue(allChecked, "soucis select all, tous les fligths ne sont pas selectionnée");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_select()
        {

            //prepare
            DateTime dateTest = DateTime.Now;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            var modalDelete = flightPage.ClickMassiveDelete();

            modalDelete.SetDateFromAndTo(dateTest, dateTest);
            //select site to filter 
            modalDelete.SelectSiteToFilter();
            modalDelete.ClickSearchButton();

            var firstFlightNumber = modalDelete.GetFirstFlightNumber();
            modalDelete.SelectFirstFlight();
            var lastnumber = modalDelete.GetLastFlightNumber();
            modalDelete.SelectLastFlight();
            modalDelete.Delete();
            flightPage.Filter(FilterType.SearchFlight, firstFlightNumber);
            var result = flightPage.CheckTotalNumber();

            flightPage.Filter(FilterType.SearchFlight, lastnumber);
            var result2 = flightPage.CheckTotalNumber();

            var finalResult = result + result2;
            Assert.IsTrue(finalResult == 0, "selected flights are not deleted");

        }


        [TestMethod]
        public void Detail_internal_flight_remarks()
        {
            //prepare
            var theRemarks = "ceci est un Internal Flight Remarks";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            var detailsModal = flightPage.EditFirstFlight("");
            detailsModal.UnfoldInternal_Flg_Remarks(theRemarks);
            var detailFlg = detailsModal.EditFlightInModal();
            var remarksFromEditFlg = detailFlg.getFlightRemarks();

            //Assert
            Assert.IsTrue(theRemarks.Equals(remarksFromEditFlg), "les remarks n'ont pas les mêmes valeur");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PrintCheck_DeliveryNoteIncludeDetails()
        {
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string serviceNameIncludeZeroQuantity = TestContext.Properties["ServiceBob"].ToString();
            string serviceNameWithQuantity = TestContext.Properties["ServiceNameLP"].ToString();
            bool versionprint = true;
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Delivery Note";
            string DocFileNameZipBegin = "All_files_";
            var flightNumber = new Random().Next().ToString();
            string deliveryNote = "New2";

            // Arrange 
            HomePage homePage = LogInAsAdmin();

            var paramPage = homePage.GoToParameters_GlobalSettings();
            var paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();

            var flightPage = new FlightPage(WebDriver, TestContext);
            flightPage.RemoveAllFromSelectedDeliveryNoteReport();
            #region refresh
            homePage.Navigate();
            paramPage = homePage.GoToParameters_GlobalSettings();
            paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            #endregion

            flightPage.SelectDeliveryNoteReport(deliveryNote);
            paramFlights.UnClickPrintSafetyNote();

            flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ClearDownloads();
            flightPage.ResetFilter();
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType("FC");
            editPage.AddService(serviceNameIncludeZeroQuantity);
            editPage.SetFinalQty("0");
            editPage.WaitPageLoading();
            editPage.AddService(serviceNameWithQuantity);
            editPage.WaitPageLoading();
            flightPage = editPage.CloseViewDetails();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.SetNewState("V");
            // Print delivery note
            var reportPageNoInclude = flightPage.PrintReport(PrintType.DeliveryNote, versionprint);
            bool isPrintDeveliryNoteNoInclude = reportPageNoInclude.IsReportGenerated();
            reportPageNoInclude.Close();
            Assert.IsTrue(isPrintDeveliryNoteNoInclude, "Le rapport n'a pas été généré.");
            reportPageNoInclude.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouveNoInclude = reportPageNoInclude.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo filePdfNoInclude = new FileInfo(trouveNoInclude);
            filePdfNoInclude.Refresh();
            Assert.IsTrue(filePdfNoInclude.Exists, trouveNoInclude + " non généré");
            //Le PDF imprimé devrait contenir:
            PdfDocument document = PdfDocument.Open(filePdfNoInclude.FullName);
            string mots = "";
            foreach (Page p in document.GetPages())
            {
                mots = mots + p.Text.Replace(" ", "");
            }

            string[] wordsInNameWithQuantity = serviceNameWithQuantity.Split(' ');
            string[] wordsInNameWithZeroQuantity = serviceNameIncludeZeroQuantity.Split(' ');

            var count1 = mots.Split(new[] { wordsInNameWithQuantity[0] }, StringSplitOptions.None).Length - 1;
            var count2 = mots.Split(new[] { wordsInNameWithQuantity[1] }, StringSplitOptions.None).Length - 1;
            var count3 = mots.Split(new[] { wordsInNameWithQuantity[2] }, StringSplitOptions.None).Length - 1;
            Assert.AreEqual(1, count1, " les services avec une quantité different de zero n apparait pas dans le rapport");
            Assert.AreEqual(1, count2, "les services avec une quantité different de zero n apparait pas dans le rapport");
            Assert.AreEqual(2, count3, "les services avec une quantité different de zero n apparait pas dans le rapport");

            //generation du deuxieme rapport
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.ClearDownloads();
            flightPage.WaiForLoad();

            var reportPageInclude = flightPage.PrintReport(PrintType.DeliveryNote, versionprint, true);
            bool isPrintDeveliryNoteInclude = reportPageInclude.IsReportGenerated();
            reportPageInclude.Close();
            Assert.IsTrue(isPrintDeveliryNoteInclude, "Le rapport n'a pas été généré.");

            reportPageInclude.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            var trouveInclud = reportPageInclude.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);

            FileInfo filePdfInclud = new FileInfo(trouveInclud);
            filePdfInclud.Refresh();
            Assert.IsTrue(filePdfInclud.Exists, trouveInclud + " non généré");
            //Le PDF imprimé devrait contenir:
            PdfDocument documentInclud = PdfDocument.Open(filePdfInclud.FullName);
            string motsInclud = "";
            foreach (Page p in documentInclud.GetPages())
            {
                motsInclud = motsInclud + p.Text.Replace(" ", "");
            }

            //Assert
            count1 = mots.Split(new[] { wordsInNameWithQuantity[0] }, StringSplitOptions.None).Length - 1;
            count2 = mots.Split(new[] { wordsInNameWithQuantity[1] }, StringSplitOptions.None).Length - 1;
            count3 = mots.Split(new[] { wordsInNameWithQuantity[2] }, StringSplitOptions.None).Length - 1;
            Assert.AreEqual(1, count1, "les services avec une quantité different de zero n apparait pas dans le rapport");
            Assert.AreEqual(1, count2, "les services avec une quantité different de zero n apparait pas dans le rapport");
            Assert.AreEqual(2, count3, "les services avec une quantité different de zero n apparait pas dans le rapport");

            // Service=> Service2- AIR FRESHNER TVS + Service for BOB
            motsInclud = motsInclud.Replace("Service2", "");

            count1 = motsInclud.Split(new[] { wordsInNameWithZeroQuantity[0] }, StringSplitOptions.None).Length - 1;
            count2 = motsInclud.Split(new[] { wordsInNameWithZeroQuantity[1] }, StringSplitOptions.None).Length - 1;
            count3 = motsInclud.Split(new[] { wordsInNameWithZeroQuantity[2] }, StringSplitOptions.None).Length - 1;
            
            Assert.AreEqual(1, count1, "les services avec une quantité zero n apparait pas dans le rapport pourtant Include details with no quantity est cochée");
            Assert.AreEqual(1, count2, "les services avec une quantité zero n apparait pas dans le rapport pourtant Include details with no quantity est cochée");
            Assert.AreEqual(1, count3, "les services avec une quantité zero n apparait pas dans le rapport pourtant Include details with no quantity est cochée");
        }


        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PrintCheck_DeliveryNoteNew()
        {
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string serviceNameNoInvoiced = TestContext.Properties["ServiceBob"].ToString();
            string serviceNameInvoiced = TestContext.Properties["ServiceNameLP"].ToString();

            string valid = "Valid";

            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Delivery Note";
            string DocFileNameZipBegin = "All_files_";

            var deliveryNote = "New";
            bool versionprint = true;
            var flightNumber = new Random().Next().ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            //information Site
            homePage.Navigate();
            var settingsSitesPage = homePage.GoToParameters_Sites();
            settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, siteFrom);
            settingsSitesPage.ClickOnFirstSite();
            settingsSitesPage.ClickToInformations();
            var adress = settingsSitesPage.GetAddress().Replace(" ", "");
            var adress2 = settingsSitesPage.GetAddress2().Replace(" ", "");
            var city = settingsSitesPage.GetCity().Replace(" ", "");
            var zipcode = settingsSitesPage.GetZipCode().Replace(" ", "");

            //Information Counters 
            settingsSitesPage.ClickToCounters();
            List<string> counters = settingsSitesPage.GetAllCounterst();

            //Information Customer 
            homePage.Navigate();
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerBob);
            var customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            var customername = customerGeneralInformationsPage.GetCustomerName();

            var contactcustomer = customerGeneralInformationsPage.GoToContactsPage();
            contactcustomer.ClickFirstContact();
            var namecustomer = contactcustomer.GetContactName();
            var contactadress = contactcustomer.GetContactAdress();

            var paramPage = homePage.GoToParameters_GlobalSettings();
            var paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            var flightPage = new FlightPage(WebDriver, TestContext);
            flightPage.RemoveAllFromSelectedDeliveryNoteReport();
            #region refresh
            homePage.Navigate();
            paramPage = homePage.GoToParameters_GlobalSettings();
            paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            flightPage.SelectDeliveryNoteReport(deliveryNote);
            #endregion

            var servicePage = new ServicePage(WebDriver, TestContext).GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameNoInvoiced);
            var pricePage = servicePage.ClickOnFirstService();

            var generalInfo = pricePage.ClickOnGeneralInformationTab();
            if (generalInfo.IsInvoiced())
            {
                generalInfo.SetNoInvoised();
            }
            servicePage = generalInfo.BackToList();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameInvoiced);
            pricePage = servicePage.ClickOnFirstService();
            generalInfo = pricePage.ClickOnGeneralInformationTab();
            if (!generalInfo.IsInvoiced())
            {
                generalInfo.SetNoInvoised();
            }
            flightPage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType("BOB");
            editPage.AddService(serviceNameNoInvoiced);
            //editPage.WaitPageLoading();
            editPage.AddService(serviceNameInvoiced);
            editPage.SetFinalQty("11");
            //editPage.WaitPageLoading();

            var fligthnumber = editPage.GetFlightNumber();
            var LegName = editPage.GetLegName();
            var deliverynum = editPage.GetdeliveryNum();
            var customer_numéro_vol_date = (customername + " : " + flightNumber + " - " + DateTime.Now.ToString("dd/MM/yyyy")).Replace(" ", "");

            flightPage = editPage.CloseViewDetails();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.SetNewState("V");
            flightPage.ShowStatusFilter();
            flightPage.Filter(FilterType.Status, valid);

            var nbflightvalid = flightPage.CheckTotalNumber();

            //1- le filtre par status
            Assert.AreEqual(1, nbflightvalid, "Le filtre ne prend pas en compte le statut valid.");

            if (flightPage.CheckTotalNumber() > 0)
            {
                flightPage.ClearDownloads();
                var reportPageNoInclude = flightPage.PrintReport(PrintType.DeliveryNote, versionprint);
                bool isPrintDeveliryNoteNoInclude = reportPageNoInclude.IsReportGenerated();
                reportPageNoInclude.Close();
                Assert.IsTrue(isPrintDeveliryNoteNoInclude, "Le rapport n'a pas été généré.");
                reportPageNoInclude.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouveNoInclude = reportPageNoInclude.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
                FileInfo filePdfNoInclude = new FileInfo(trouveNoInclude);
                filePdfNoInclude.Refresh();
                Assert.IsTrue(filePdfNoInclude.Exists, trouveNoInclude + " non généré");
                PdfDocument document = PdfDocument.Open(filePdfNoInclude.FullName);
                string mots = "";
                foreach (Page p in document.GetPages())
                {
                    mots += p.Text;
                }

                string ch = mots.Replace(" ", "");

                // 2- le logo du site

                Page page1 = document.GetPage(1);
                List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
                Assert.AreEqual(1, images.Count, "Pas d'image dans le PDF");
                IPdfImage image = images[0];
                FileInfo fiPdf = new FileInfo(downloadsPath + "\\pizza_test.png");
                if (fiPdf.Exists)
                {
                    fiPdf.Delete();
                }
                File.WriteAllBytes(downloadsPath + "\\pizza_test.png", image.RawBytes.ToArray<byte>());
                fiPdf.Refresh();
                Assert.IsTrue(fiPdf.Exists, "fichier non trouvé");
                FileInfo fiTest;
                fiTest = new FileInfo(TestContext.TestDeploymentDir + "\\pizza_petite.png");
                Assert.IsTrue(fiTest.Exists, "fichier non trouvé (cas 2)");

                // 3.verifier que les informations générales du site (nom, adresse…) sont présents

                Assert.IsTrue(ch.Contains(siteFrom) && ch.Contains(adress) && ch.Contains(adress2) && ch.Contains(city) && ch.Contains(zipcode), " les informations générales du site (nom, adresse…) ne sont pas présents dans le PDF");

                // 4.verifier que le nom du customer est présent, avec le numéro de vol et la date

                Assert.IsTrue(ch.Contains(customer_numéro_vol_date), " le nom du customer avec le numéro de vol et la date ne sont pas présents dans le PDF");

                // 5.Verfier le leg (route) apparait bien

                //Leg 1 : ACE - BCN
                //1670743467
                string[] LegNameSplit = LegName.Split('\n');
                string[] LegNameSplit2 = LegNameSplit[0].Split(':');
                //ACE - BCN
                Assert.IsTrue(ch.Contains(LegNameSplit2[0].Trim()), " le leg n'apparait pas bien dans le PDF");
                //1670743467
                Assert.IsTrue(ch.Contains(LegNameSplit[0]), " le leg n'apparait pas bien dans le PDF");

                // 6.verifier que les services sont listés

                Assert.IsTrue(ch.Contains(serviceNameNoInvoiced.Replace(" ", "")), " les service « is invoice »  leg n'apparait pas bien dans le PDF");

                // 7.verifier que le délivery note n° est bien généré avec le flightid du vol

                Assert.IsTrue(mots.Contains(deliverynum), " les service « is invoice »  leg n'apparait pas bien dans le PDF");

                // 8. Vérifier que le customs n° est généré depuis setting/site/counters
                bool trouveNumero = false;
                foreach (var c in counters)
                {
                    string numero = (int.Parse(c) + 1).ToString();
                    int count = mots.Split(new[] { numero }, StringSplitOptions.None).Length;
                    if (count >= 2)
                    {
                        trouveNumero = true;
                        break;
                    }
                }
                Assert.IsTrue(trouveNumero, " le customs n° est généré depuis setting/site/counters ne sont pas présents dans le PDF");


                // 9. Vérifier que les quantités des services sont correctes (somme du nombre de services présents dans le vol)

                Assert.IsTrue(2 + 1 <= mots.Split(new[] { "11" }, StringSplitOptions.None).Length, " les quantités des services ne sont pas correctes dans le PDF");

                // 10.verifier qu’un champs « adjustments » est présent et vide

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PrintCheck_TrackSheet()
        {
            //Prepare
            string flightNameStart = "FLIGHTTRACKSHEET";
            string flightNumber = flightNameStart + "_" + new Random().Next(10, 99).ToString();
            string flightNumber1 = flightNameStart + "_" + new Random().Next(100, 999).ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool versionPrint = true;
            bool deliveryNote = false;
            bool isQuantityNull = true;
            bool isOneFlightPerPage = true;
            string DocFileNamePdfBegin = "Print TrackSheet";
            string DocFileNameZipBegin = "All_files_";
            string service = "AIR FRESHNER TVS";
            string qty = "0";
            string stateToSet = "V";
            string valid = "Valid";

            //Arrange
            var homePage = LogInAsAdmin();

            //act
            homePage.PurgeDownloads();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();

            //Create Flight 1 
            var flightCreatePageModal = flightPage.FlightCreatePage();
            flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);
            var flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
            flightDetailsPage.AddGuestType();
            flightDetailsPage.AddService(service);
            flightDetailsPage.SetFinalQtyServiceFlight(qty);
            flightDetailsPage.CloseViewDetails();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.SetNewState(stateToSet);
            flightPage.ResetFilter();

            //Create Flight 2
            flightCreatePageModal = flightPage.FlightCreatePage();
            flightCreatePageModal.FillField_CreatNewFlight(flightNumber1, customerBob, aircraft, siteFrom, siteTo);
            flightDetailsPage = flightPage.EditFirstFlight(flightNumber1);
            flightDetailsPage.AddGuestType();
            flightDetailsPage.AddService(service);
            flightDetailsPage.SetFinalQtyServiceFlight(qty);
            flightDetailsPage.CloseViewDetails();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber1);
            flightPage.SetNewState(stateToSet);
            flightPage.ResetFilter();

            flightPage.ShowStatusFilter();
            flightPage.Filter(FilterType.Status, valid);
            flightPage.Filter(FilterType.SearchFlight, flightNameStart);
            var nbrOfFlights = flightPage.CheckTotalNumber();
            try
            {
                flightPage.ClearDownloads();
                var trackSheetPrintReport = flightPage.PrintReport(PrintType.TrackSheet, versionPrint, deliveryNote, isOneFlightPerPage, isQuantityNull);
                trackSheetPrintReport.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                trackSheetPrintReport.Close();
                string trouve = trackSheetPrintReport.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fichier = new FileInfo(trouve);
                fichier.Refresh();
                Assert.IsTrue(fichier.Exists, "Pas de print pdf");
                //Le PDF imprimé devrait contenir:
                PdfDocument document = PdfDocument.Open(fichier.FullName);
                List<string> mots = new List<string>();
                foreach (Page p in document.GetPages())
                {
                    foreach (var mot in p.GetWords())
                    {
                        mots.Add(mot.Text);
                    }
                }
                int nbrOfQtNull = 0;
                foreach (string mot in mots)
                {
                    if (mot == "0")
                    {
                        nbrOfQtNull++;
                    }
                }
                bool isAfficheFlightValid = flightPage.AfficheFlightValid(mots, flightNumber, flightNumber1);
                var documentNbrOfPages = document.NumberOfPages;

                //Assert
                Assert.IsTrue(isAfficheFlightValid, "Seuls les vols en statut V ressortent");
                Assert.IsTrue(nbrOfQtNull >= 1, "l'affichage de la quantité nulle n'est pas fonctionnel");
                Assert.AreEqual(documentNbrOfPages, nbrOfFlights, "un vol par page n'est pas fonctionnel !");
            }
            finally
            {
                //Delete Flights
                FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();
                massiveDeleteModal.SetFlightName(flightNameStart);
                massiveDeleteModal.ClickSearchButton();
                massiveDeleteModal.ClickSelectAllButton();
                massiveDeleteModal.Delete();

            }

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PrintCheck_DeliveryNoteINotGroupByCategory()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string flightNumber = new Random().Next().ToString();
            string serviceCategorie = TestContext.Properties["Prodman_Needs_ServiceCategory1"].ToString();
            string guestTypeBob = TestContext.Properties["GuestNameBob"].ToString();
            var serviceCategory = string.Empty;
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool versionPrint = true;
            bool deliveryNote = false;
            bool groupedBy = true;
            string DocFileNamePdfBegin = "Print Delivery Note";
            string DocFileNameZipBegin = "All_files_";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

            //Edit the flight
            var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.AddGuestType("BOB");
            editPage.AddService(serviceName);

            flightPage = editPage.CloseViewDetails();

            homePage.Navigate();
            var servicesPage = homePage.GoToCustomers_ServicePage();
            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);
            if (servicesPage.CheckTotalNumber() > 0)
            {
                var servicePricePage = servicesPage.ClickOnFirstService();
                ServiceGeneralInformationPage serviceGeneralInformationPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceCategory = serviceGeneralInformationPage.GetCategory();
            }

            homePage.Navigate();
            flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);

            if (flightPage.CheckTotalNumber() > 0)
            {
                flightPage.ClearDownloads();
                flightPage.PurgeDownloads();
                var reportPageGroupedByCategory = flightPage.PrintReport(PrintType.DeliveryNote, versionPrint);
                bool isPrintDeveliryNote = reportPageGroupedByCategory.IsReportGenerated();
                reportPageGroupedByCategory.Close();
                reportPageGroupedByCategory.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPageGroupedByCategory.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
                FileInfo filePdf = new FileInfo(trouve);
                filePdf.Refresh();
                Assert.IsTrue(filePdf.Exists, trouve + " non généré");
                PdfDocument document = PdfDocument.Open(filePdf.FullName);
                List<string> mots = new List<string>();
                foreach (Page page in document.GetPages())
                {
                    mots.AddRange(page.GetWords().Select(m => m.Text));
                }
                Assert.IsTrue(mots.Contains(serviceCategory), "Le rapport ne se trie pas par catégory de service");
            }
            if (flightPage.CheckTotalNumber() > 0)
            {
                flightPage.ClearDownloads();
                flightPage.PurgeDownloads();
                var reportPageGroupedByCategory = flightPage.PrintReport(PrintType.DeliveryNote, versionPrint, deliveryNote, false, false, groupedBy);
                bool isPrintDeveliryNote = reportPageGroupedByCategory.IsReportGenerated();
                reportPageGroupedByCategory.Close();
                reportPageGroupedByCategory.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPageGroupedByCategory.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
                FileInfo filePdf = new FileInfo(trouve);
                filePdf.Refresh();
                Assert.IsTrue(filePdf.Exists, trouve + " non généré");
                PdfDocument document = PdfDocument.Open(filePdf.FullName);
                List<string> mots = new List<string>();
                foreach (Page page in document.GetPages())
                {
                    mots.AddRange(page.GetWords().Select(m => m.Text));
                }
                Assert.IsTrue(mots.Contains(guestTypeBob), "Le rapport ne se trie pas par guest type");
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PrintCheck_DeliveryNoteIncludeSafetyNote()
        {
            string flightNumber = new Random().Next() + "_" + DateTime.Now.ToString("yyyy-MM-dd");
            string deliveryNoteNew2 = "New2";

            // Arrange 
            HomePage homePage = LogInAsAdmin();
            homePage.ClearDownloads();
            homePage.PurgeDownloads();
            //try
            //{
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Delivery Note";
            string DocFileNameZipBegin = "All_files_";

            var paramPage = homePage.GoToParameters_GlobalSettings();
            var paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();

            var flightPage = new FlightPage(WebDriver, TestContext);
            flightPage.RemoveAllFromSelectedDeliveryNoteReport();
            #region refresh
            homePage.Navigate();
            paramPage = homePage.GoToParameters_GlobalSettings();
            paramFlights = paramPage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            #endregion

            flightPage.SelectDeliveryNoteReport(deliveryNoteNew2);
            paramFlights.ClickPrintSafetyNote();

            var siteSettings = homePage.GoToParameters_Sites();
            siteSettings.ClickOnFirstSite();
            siteSettings.ClickToInformations();
            var agreementNumber = siteSettings.GetAirportAgreement();
            var safetyNote = siteSettings.GetSafetyNote();
            var healthApproval = siteSettings.GetHealthApprovalNumber();
            siteSettings.ClickToCounters();
            var customsNumber = siteSettings.GetCustomNumber();

            flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, "ACE");
            FlightCreateModalPage newFlight = flightPage.FlightCreatePage();
            newFlight.FillField_CreatNewFlight(flightNumber, null, "B717", "ACE", "AAL");
            FlightDetailsPage flightDetails = flightPage.EditFirstFlight(flightNumber);
            flightDetails.AddGuestType("FC");
            flightDetails.AddService("testServiceForTest");
            flightDetails.CloseViewDetails();

            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, "ACE");

            flightPage.Filter(FilterType.SearchFlight, flightNumber);

            flightPage.ShowStatusFilter();
            flightPage.Filter(FilterType.Status, "Valid");
            if (flightPage.CheckTotalNumber() == 0)
            {
                flightPage.Filter(FilterType.Status, "Preval");
                flightPage.SetNewState("V");
                flightPage.Filter(FilterType.Status, "Valid");
            }
            var flightDate = flightPage.GetDateState().Split(' ')[0];
            var flightETD = flightPage.GetETD();

            var editFlight = flightPage.EditFirstFlight(flightNumber);
            var deliverynum = editFlight.GetdeliveryNum();
            flightPage = editFlight.CloseViewDetails();

            flightPage.Filter(FilterType.SearchFlight, flightNumber);
            var deliveryNote = flightPage.PrintReport(PrintType.DeliveryNote, true, true, false, false, false, true);
            bool isPrintDeliveryNote = deliveryNote.IsReportGenerated();
            deliveryNote.Close();
            Assert.IsTrue(isPrintDeliveryNote, "Le rapport n'a pas été généré.");
            deliveryNote.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = deliveryNote.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo filePdfNoInclude = new FileInfo(trouve);
            filePdfNoInclude.Refresh();
            Assert.IsTrue(filePdfNoInclude.Exists, trouve + " non généré");
            PdfDocument document = PdfDocument.Open(filePdfNoInclude.FullName);
            //List<string> mots = new List<string>();
            string mots = "";
            foreach (Page p in document.GetPages())
            {
                //mots.AddRange(p.GetWords().Select(m => m.Text));
                mots += p.Text.Replace(" ", "");
            }

            //SafeMode activé
            Assert.IsTrue(mots.Contains(healthApproval), healthApproval + " pas trouvé");
            var footer = safetyNote.Split(' ');
            Assert.IsTrue(mots.Contains(footer[footer.Length - 1]), footer[footer.Length - 1] + " pas trouvé");
            //customsNumber - 1
            int customerNumberInt = int.Parse(customsNumber);
            // compte à partir de zéro - pas bon pour le parallélisme
            customerNumberInt += 1;
            string customerNumberString = customerNumberInt.ToString();
            Assert.IsTrue(mots.Contains(customerNumberString), customerNumberString + " pas trouvé");
            Assert.IsTrue(mots.Contains(agreementNumber), agreementNumber + " pas trouvé");
            Assert.IsTrue(mots.Contains(flightNumber), flightNumber + " pas trouvé");
            Assert.IsTrue(mots.Contains(flightDate), flightDate + " pas trouvé");
            Assert.IsTrue(mots.Contains(flightETD), flightETD + " pas trouvé");
            Assert.IsTrue(mots.Contains(deliverynum), deliverynum + " pas trouvé");

            // Count via Split
            int count = mots.Split(new[] { "Trolley" }, StringSplitOptions.None).Length;
            // 1 => 2 length
            // 2 => 3 length
            // 3 => 4 length
            Assert.IsTrue(count >= 18, count + " pas plus que 18");

            //}
            //finally
            //{
            //var 
            paramFlights = homePage.GoToParametres_Flights();
            paramFlights.GoToDeliveryReport();
            paramFlights.UnClickPrintSafetyNote();
            //}

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_SiteFilter()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FilterType.Sites, "ACE");
            var numberOfFlightsACE = flightPage.CheckTotalNumber();
            flightPage.Filter(FilterType.Sites, "MAD");
            var numberOfFlightsMad = flightPage.CheckTotalNumber();
            Assert.AreNotEqual(numberOfFlightsMad, numberOfFlightsACE, "Le filter par Site n'est pas fonctionnel");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_ResetFilter()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FilterType.SearchFlight, "Reset Test Search");
            flightPage.ResetFilter();
            string searchValue = flightPage.GetSearchValue();
            Assert.IsTrue(searchValue == string.Empty);

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_DockNumberFilter()
        {
            string dockNumberToFilter = "192";
            
            var homePage = LogInAsAdmin();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Status, "Preval");
            // si dans l'index on a un dock number, alors après le filtre by dock number fonctionne
            flightPage.SettDockNumber(dockNumberToFilter);

            flightPage.ScrollUntilDockNumberFilterIsVisible();
            flightPage.Filter(FlightPage.FilterType.DockNumber, dockNumberToFilter);
            if (flightPage.CheckTotalNumber() == 0)
            {
                Assert.Fail("Aucune Flight ne correspond à ce filtrage par Dock Number");
            }
            Assert.IsTrue(flightPage.CheckTotalNumber() == flightPage.GetDockNumbers().Count() && flightPage.GetDockNumbers().All(dock => dock == dockNumberToFilter), "L’application de filtre par Dock Number ne fonctionne pas correctement");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_DriverFilter()
        {
            string driverToFilter = "276";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            // si dans l'index on a un driver, alors après le filtre by driver fonctionne
            flightPage.SetDriver(driverToFilter);

            flightPage.Filter(FlightPage.FilterType.Driver, driverToFilter);

            if (flightPage.CheckTotalNumber() == 0)
            {
                Assert.Fail("Aucune données ne correspond à ce filtrage par le driver");
            }
            Assert.IsTrue(flightPage.CheckTotalNumber() == flightPage.GetDrivers().Select(d => d.Text.Trim()).ToList().Count() && flightPage.GetDrivers().Select(d => d.Text.Trim()).ToList().All(driver => driver == driverToFilter), "L’application de filtre sur les driver ne fonctionne pas correctement");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Massivedelete_UseCase()
        {
            Random rand = new Random();
            DateTime dateFin = DateTime.Today;
            DateTime dateDebut = dateFin.AddYears(-3);

            int range = (DateTime.Today - dateDebut.AddYears(1)).Days;
            dateDebut = dateDebut.AddDays(rand.Next(range));

            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();

            FlightMassiveDeleteModalPage massiveDeleteModal = flightPage.ClickMassiveDelete();
            massiveDeleteModal.SetDateFromAndTo(dateDebut, dateFin);
            massiveDeleteModal.ClickSearchButton();

            string date = massiveDeleteModal.GetFirstRowDate();
            string site = massiveDeleteModal.GetFirstRowSite();
            massiveDeleteModal.ClickFirstUseCase();
            string legsNum = massiveDeleteModal.GetLegsNumber();

            Assert.IsTrue(date != null && site != null, "Aucun résultat de Flight trouvé pour une recherche de massive delete.");

            bool linkIsClicked = massiveDeleteModal.ClickOnResultLinkAndCheckFlight(date, site);
            string legsNumFlight = massiveDeleteModal.GetLegsNumberFromFlight();

            Assert.IsTrue(linkIsClicked && legsNum.Equals(legsNum), "Les UseCase affichés n'est pas corrects.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexStatePrevalidated()
        {
            string site = TestContext.Properties["SiteACE"].ToString();

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, site);
            var flightNo = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNo);
            var previousPrevalStatus = flightPage.GetFirstFlightStatus("P");
            var previousValidStatus = flightPage.GetFirstFlightStatus("V");
            var previousInvoiceStatus = flightPage.GetFirstFlightStatus("I");
            try
            {
                flightPage.SetNewState("P");
                var newPrevalStatus = flightPage.GetFirstFlightStatus("P");
                var isNotChanged = previousPrevalStatus.Equals(newPrevalStatus);
                Assert.IsFalse(isNotChanged, "Preval statut n'a pas changé");
            }
            finally
            {
                if (previousPrevalStatus != flightPage.GetFirstFlightStatus("P"))
                {
                    flightPage.SetNewState("P");
                }
                if (previousValidStatus != flightPage.GetFirstFlightStatus("V"))
                {
                    flightPage.SetNewState("V");
                }
                if (previousInvoiceStatus != flightPage.GetFirstFlightStatus("I"))
                {
                    flightPage.SetNewState("I");
                }
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexFlightNumber()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, "ACE");
            var flightNo = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNo);
            var previousValidStatus = flightPage.GetFirstFlightStatus("V");
            var randomNumber1 = new Random().Next(10);
            var newFlightNumber = flightPage.GetFirstFlightNumber() + randomNumber1.ToString();
            try
            {
                if (previousValidStatus)
                {
                    flightPage.SetNewState("V");
                }

                flightPage.ModifyFlightName(newFlightNumber);
                var newPrevalStatus = flightPage.GetFirstFlightStatus("P");
                WebDriver.Navigate().Refresh();
                var actualFlightnumber = flightPage.GetFirstFlightNumber();
                Assert.IsTrue(!flightNo.Equals(actualFlightnumber), "Numero vol n'a pas changé");
            }
            finally
            {
                flightPage.Filter(FilterType.SearchFlight, newFlightNumber);

                flightPage.ModifyFlightName(flightNo);

                if (previousValidStatus != flightPage.GetFirstFlightStatus("V"))
                {
                    flightPage.SetNewState("V");
                }
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexStateInvoice()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.PageSize("100");
            var list = flightPage.GetFlightList();
            var index = flightPage.GetFirstIndexInvoiceAllowed();
            Assert.AreNotEqual(index, -1, "there are no flight possible to test state invoice");
            var test1 = flightPage.GetState(index);
            if (test1 == "V")
            {
                flightPage.ChangeStateInvoice(index, 3);
            }
            else
            {
                flightPage.ChangeStateInvoice(index, 2);
                flightPage.ChangeStateInvoice(index, 3);
            }
            homePage.GoToFlights_FlightPage();
            var test2 = flightPage.GetState(index);
            // rejouer
            flightPage.ChangeStateInvoice(index, 3);
            flightPage.ChangeStateInvoice(index, 2);
            Assert.AreNotEqual(test1, test2, "test index invoice has failed");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexAC()
        {
            //Prepare
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string oldAircraft = "B717";
            string newAircraft = TestContext.Properties["Aircraft"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //act
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, oldAircraft, siteFrom, siteTo);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                string oldvalue = flightPage.GetIndexAirCraftValue();
                flightPage.CliCkOnFirstFlight(2);
                flightPage.EditAircraftAC(newAircraft);
                flightPage.WaitPageLoading();
                WebDriver.Navigate().Refresh();
                string modifiedvalue = flightPage.GetIndexAirCraftValue();

                //Assert
                Assert.AreNotEqual(oldvalue, modifiedvalue, MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE);
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.SetDateFromAndTo(DateTime.Now, DateTime.Now);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_FlightAlerts()
        {
            string alertTypes1 = "Missing Dock Number";
            string modalMissingDockNumber = "MissingDockNumber";
            string customer = TestContext.Properties["CustomerDelivery"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string flightNumber = new Random().Next().ToString();
            string toSiteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string flightAlertsToFilter = "MissingDockNumber";
            
            var homePage = LogInAsAdmin();
            //Act
            try
            {
                var settingsFlightsPage = homePage.GoToParametres_Flights();
                settingsFlightsPage.GoToFlightAlerts();
                settingsFlightsPage.SelectAlertTypes(alertTypes1);
                settingsFlightsPage.AddCustomerSubscriptionToAlertTypes(modalMissingDockNumber, customer, siteACE);
                homePage = new HomePage(WebDriver, TestContext);
                homePage.Navigate();
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                var flightCreatePage = flightPage.FlightCreatePage();
                // Add New Flight without  Dock Number 
                flightCreatePage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteACE, toSiteBCN);
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
                flightPage.Filter(FlightPage.FilterType.FlightAlerts, flightAlertsToFilter);
                Assert.IsTrue(flightPage.IsResultFlightAlertsFiltreVisibile(alertTypes1), "L’application de filtre sur les flight alerts ne fonctionne pas correctement");
            }
            finally
            {
                var settingsFlightsPage = homePage.GoToParametres_Flights();
                settingsFlightsPage.GoToFlightAlerts();
                settingsFlightsPage.SelectAlertTypes(alertTypes1);
                settingsFlightsPage.DeleteCustomerSubscription(modalMissingDockNumber, customer, siteACE);
            }


        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexETD()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, "ACE");
            var flightNo = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNo);
            var previousValidStatus = flightPage.GetFirstFlightStatus("V");
            int rndHour = new Random().Next(0, 12);
            int rndMinute = new Random().Next(10, 59);
            var random = rndHour.ToString();
            var previousETD = flightPage.GetETDValue();
            try
            {
                if (previousValidStatus)
                {
                    flightPage.SetNewState("V");
                }

                flightPage.SetETD(random);
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNo);

                var newETDValue = flightPage.GetETDValue();
                Assert.IsTrue(!previousETD.Equals(newETDValue), "ETD n'a pas changé");
            }
            finally
            {
                flightPage.SetETD(previousETD);

                if (previousValidStatus != flightPage.GetFirstFlightStatus("V"))
                {
                    flightPage.SetNewState("V");
                }
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexArrNumber()
        {
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            Random random = new Random();
            string randomIndexDockNumber = random.Next(10000, 100000).ToString();

            var homePage = LogInAsAdmin();
            // Act Flight
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, "-", siteFrom, siteTo);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                string oldarrnumber = flightPage.GetIndexArrNumberValue();
                string safeOldArrNumber = string.IsNullOrEmpty(oldarrnumber) ? "-" : oldarrnumber;
                flightPage.SetArrNumberValue(randomIndexDockNumber);
                WebDriver.Navigate().Refresh();
                string modifiedarrnumber = flightPage.GetIndexArrNumberValue();
                // Assert
                Assert.AreNotEqual(safeOldArrNumber, modifiedarrnumber, MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE);
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexDockNumber()
        {
            // Prepare
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            Random random = new Random();
            int randomIndexDockNumber = random.Next(10000, 100000);
            string randomIndexDockNumberStr = randomIndexDockNumber.ToString();

            var homePage = LogInAsAdmin();
            // Act Flight
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, "-", siteFrom, siteTo);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                string oldarrnumber = flightPage.GetIndexDockNumberValue();
                string safeOldArrNumber = string.IsNullOrEmpty(oldarrnumber) ? "-" : oldarrnumber;
                flightPage.SetIndexDockNumberValue(randomIndexDockNumberStr);
                WebDriver.Navigate().Refresh();
                string modifiedarrnumber = flightPage.GetIndexDockNumberValue();
                // Assert
                Assert.AreNotEqual(safeOldArrNumber, modifiedarrnumber, MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE);
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_BarsetStockExport()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersion = false;
            var message = "Select at least one flight status";
            string status = "Valid";
            bool export = true;
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ClearDownloads();
            flightPage.ResetFilter();
            flightPage.ExportFlights(ExportType.BarsetStockExport, newVersion, status, export);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = flightPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Le fichier exporté ne se trouve pas dans le dossier de téléchargements de WinRest.");
            flightPage.ExportFlights(ExportType.BarsetStockExport, newVersion, string.Empty, export);
            var isMessageDisplayed = flightPage.GetMessagebarsetstock().Equals(message);
            Assert.IsTrue(isMessageDisplayed, "si aucun statut de flight n’est sélectionné, l'exportation ne devrait pas être possible et un message d'erreur devrait apparaître.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexRegist()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            // ne marche pas avec les Valid (ligne non éditabe)
            flightPage.Filter(FilterType.Status, "Preval");
            flightPage.PageSize("100");
            var flightCount = flightPage.GetFlightList().Count;
            Random random = new Random();
            int randomIndex = random.Next(2, flightCount);
            string oldindexRegist = flightPage.GetIndexRegist(randomIndex);
            string randomIndexRegist = random.Next(2, 1000).ToString();
            if (randomIndexRegist != oldindexRegist)
            {
                flightPage.ChangeIndexRegist(randomIndex, randomIndexRegist);
            }
            else
            {
                flightPage.ChangeIndexRegist(randomIndex, randomIndexRegist + 1);
            }
            homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Status, "Preval");
            flightPage.PageSize("100");
            string newindexRegist = flightPage.GetIndexRegist(randomIndex);
            Assert.AreNotEqual(oldindexRegist, newindexRegist, "index regist has not changed succesfully");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_FlightCounting()
        {
            //HomePage homePage = LogInAsAdmin();
            //FlightPage flightPage = homePage.GoToFlights_FlightPage();
            //flightPage.ResetFilter();
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_TotalFlightIndex()
        {
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string customerBob = TestContext.Properties["CustomerLpFilter"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            // Act Flight
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, siteFrom);
            flightPage.SetDateState(DateUtils.Now);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, "-", siteFrom, siteTo);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            string flightsnumber = flightPage.GetIndexFlightsNumber();
            var modifiedarrnumber = flightPage.GetTotalNumberLegs();
            // Assert
            Assert.AreEqual(flightsnumber, modifiedarrnumber.ToString(), MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE);
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexETA()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string rndHour = new Random().Next(1, 13).ToString("D2");

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, site);
            var flightNo = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNo);
            var previousValidStatus = flightPage.GetFirstFlightStatus("V");
            var previousETA = flightPage.GetETAValue();
            try
            {
                if (previousValidStatus)
                {
                    flightPage.SetNewState("V");
                }

                flightPage.SetETA(rndHour);
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNo);
                var newETAValue = flightPage.GetETAValue().Substring(0, 2);
                var isEdited = rndHour.Equals(newETAValue);
                Assert.IsTrue(isEdited, "ETA n'a pas changé");
            }
            finally
            {
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNo);
                flightPage.SetETA(previousETA);
                if (previousValidStatus != flightPage.GetFirstFlightStatus("V"))
                {
                    flightPage.SetNewState("V");
                }
            }
        }
       

        [TestMethod]
       // [Timeout(_timeout)]
        public void FL_FLIG_Calendar_SortByFlightNumber()
        {
            //Prepare
            string sortBy = "Sort by flight no";

            //Arrange
            HomePage homePage = LogInAsAdmin();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();

            //Act
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();

            //Filter par num vol
            hedomadairePage.ResetFilter();
            hedomadairePage.Filter(HedomadairePage.FilterType.SortBy, sortBy);
            var isSortedByFlightNo = hedomadairePage.IsSortedByFlightNo();

            //Assert
            Assert.IsTrue(isSortedByFlightNo, MessageErreur.FILTRE_ERRONE, sortBy);
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_HedomadaireViewOP245()
        {
            //Prepare
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string ServiceNameNoSmpl = TestContext.Properties["ServiceBob"].ToString();
            string ServiceNameSmpl = TestContext.Properties["TrolleyService"].ToString();
            bool newVersion = true;
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            DateTime now = DateTime.Now;

            //Arrange
            HomePage homePage = LogInAsAdmin();
           
            //Act
            //Verifie si le customer est Buy on board
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customer);
            var customerGeneralInformationsPage = customersPage.SelectFirstCustomer();
            var valueairvision = customerGeneralInformationsPage.GetCustomerAirVision();
            Assert.IsTrue(valueairvision.Equals("Yes"), "le customer n'est pas activé sur airvision");
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ServiceNameSmpl);
            var pricePage = servicePage.ClickOnFirstService();
            pricePage.WaitPageLoading();
            var generalInfo = pricePage.ClickOnGeneralInformationTab();
            try
            {
                if (!generalInfo.IsSPML())
                {
                    generalInfo.SetSPML();
                    generalInfo.SpecialMeals();
                    generalInfo.WaitPageLoading();
                }
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.ClearDownloads();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

                var legfirstflight = flightPage.GetFirstFlightLeg();

                var editPage = flightPage.EditFirstFlight(flightNumber);
                editPage.WaitPageLoading();
                editPage.AddGuestType("FC");
                editPage.AddService(ServiceNameNoSmpl);
                editPage.AddService(ServiceNameSmpl);
                editPage.WaitPageLoading();
                editPage.CloseViewDetails();
                flightPage.SetNewState("V");
                flightPage.Refrech(); 
                HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
                hedomadairePage.ResetFilter();

                hedomadairePage.ClearDownloads();
                DeleteAllFileDownload();

                hedomadairePage.Filter(HedomadairePage.FilterType.Search, flightNumber);
                editPage.WaitPageLoading();
                string flightNumberHebdo = hedomadairePage.GetFirstFlightNumber();
                Assert.AreEqual(flightNumber.Trim(), flightNumberHebdo.Trim());
                hedomadairePage.ExportHedomadaire(HedomadairePage.ExportType.OP245, newVersion);
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();
                var correctDownloadedFile = flightPage.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile);
                var fileName = correctDownloadedFile.Name;
                var filePath = System.IO.Path.Combine(downloadsPath, fileName);
                List<string> firstRowValues = OpenXmlExcel.ReadFirstRow("MASTER", correctDownloadedFile.FullName);
                //Assert
                Assert.IsTrue(firstRowValues[0].Equals("1"), "pas de 1 dans les données du fichier Excel Colonnes A");
                Assert.IsTrue(firstRowValues[1].Equals("00"), "pas de 00 dans les données du fichier Excel Colonnes B");


                string formattedDate = now.ToString("dd/MM/yyyy");
                int dayInt = int.Parse(now.ToString("dd"));
                int monthInt = int.Parse(now.ToString("MM"));
                int yearInt = int.Parse(now.ToString("yyyy"));
                Assert.IsTrue((yearInt == int.Parse(firstRowValues[2]))
                            && (monthInt == int.Parse(firstRowValues[3]))
                            && (dayInt == int.Parse(firstRowValues[4])), "date du boarding dans les colonnes CDE Incorrect.");
                Assert.IsTrue(firstRowValues[11].Equals("FC"), " Dans colonne L Guest Incorrect.");
                if (int.Parse(firstRowValues[16]) == 2)
                { Assert.IsTrue(firstRowValues[17].Equals("E"), " Dans colonne R Extra-Order Incorrect."); }
                else { Assert.IsTrue(firstRowValues[17].Equals(""), " Dans colonne R Extra-Order Incorrect."); }
                Assert.IsTrue(firstRowValues[20].Equals("Code") || firstRowValues[20].Equals(""), " La colonne U Incorrect.");
                Assert.IsTrue((firstRowValues[5].Equals(flightNumber))
                            && (firstRowValues[6].Equals(""))
                            && (firstRowValues[7].Equals(siteTo)), "infos boarding dans les colonnes FGH Incorrect.");
                Assert.IsTrue(firstRowValues[20].Equals("Code") || firstRowValues[20].Equals(""), " La colonne U Incorrect.");
                if (legfirstflight == 1)
                {
                    Assert.IsTrue((firstRowValues[8].Equals(""))
                    && (firstRowValues[9].Equals(""))
                    && (firstRowValues[10].Equals("")), " Les colonnes IGK sont Incorrect. ");
                }
            }
            finally
            {
                servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, ServiceNameSmpl);
                pricePage = servicePage.ClickOnFirstService();
                pricePage.WaitPageLoading();
                generalInfo = pricePage.ClickOnGeneralInformationTab();
                if (generalInfo.IsSPML())
                {
                    generalInfo.SetNoSPML();
                }
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexGate()
        {
            LogInAsAdmin();
            string gateNumber = new Random().Next(10, 100).ToString();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, "ACE");
            var flightNo = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNo);
            var previousValidStatus = flightPage.GetFirstFlightStatus("V");
            var previousGate = flightPage.GetGateValue();
            try
            {
                if (previousValidStatus)
                {
                    flightPage.SetNewState("V");
                    // sortir du mode Edit line
                    flightPage.ResetFilter();
                    flightPage.Filter(FilterType.SearchFlight, flightNo);
                }

                flightPage.SetGATE(gateNumber);
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNo);

                var newGateValue = flightPage.GetGateValue();
                Assert.IsTrue(!previousGate.Equals(newGateValue), "Gate n'a pas changé");
            }
            finally
            {
                flightPage.SetGATE(previousGate);

                if (previousValidStatus != flightPage.GetFirstFlightStatus("V"))
                {
                    flightPage.SetNewState("V");
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_HedomadaireViewHideCancelFlights()
        {
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            hedomadairePage.ResetFilter();
            hedomadairePage.Filter(HedomadairePage.FilterType.HideCancelledFlight, true);
            hedomadairePage.Filter(HedomadairePage.FilterType.HideCancelledFlight, true);
            Assert.IsFalse(hedomadairePage.IsFlightDisplayed(), "Le flight est affiché malgré le fait qu'il soit CANCELLED");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexDriver()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string driverValue = ((char)new Random().Next('A', 'Z' + 1)).ToString();

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, site);
            var flightNo = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNo);
            var previousValidStatus = flightPage.GetFirstFlightStatus("V");
            var previousDriver = flightPage.GetDriverValue();
            try
            {
                if (previousValidStatus)
                {
                    flightPage.SetNewState("V");
                }
                flightPage.SetDriver(driverValue);
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNo);
                var diplayedValue = flightPage.GetDriverValue();
                var isEdited = driverValue.Equals(diplayedValue);
                Assert.IsTrue(isEdited, "Driver n'a pas changé");
            }
            finally
            {
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNo);
                flightPage.SetDriver(previousDriver);

                if (previousValidStatus != flightPage.GetFirstFlightStatus("V"))
                {
                    flightPage.SetNewState("V");
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexStateValidated()
        {

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, "ACE");
            var flightNo = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNo);
            var previousPrevalStatus = flightPage.GetFirstFlightStatus("P");
            var previousValidStatus = flightPage.GetFirstFlightStatus("V");
            var previousInvoiceStatus = flightPage.GetFirstFlightStatus("I");
            try
            {
                flightPage.SetNewState("V");
                flightPage.Refrech();
                var newValidatedStatus = flightPage.GetFirstFlightStatus("V");
                Assert.IsTrue(previousValidStatus != newValidatedStatus, "Validated statut n'a pas changé");
            }
            finally
            {
                if (previousPrevalStatus != flightPage.GetFirstFlightStatus("P"))
                {
                    flightPage.SetNewState("P");
                }
                if (previousValidStatus != flightPage.GetFirstFlightStatus("V"))
                {
                    flightPage.SetNewState("V");
                }
                if (previousInvoiceStatus != flightPage.GetFirstFlightStatus("I"))
                {
                    flightPage.SetNewState("I");
                }
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexFrom()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            var flightNo = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNo);
            var previousValidStatus = flightPage.GetFirstFlightStatus("V");
            var fromValue = "ACE";

            var previousFrom = flightPage.GetFromValue();
            try
            {
                if (previousValidStatus)
                {
                    flightPage.SetNewState("V");
                }
                if (previousFrom.Equals(fromValue))
                {
                    fromValue = "MAD";
                }
                flightPage.SetFrom(fromValue);

                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNo);
                var newFromValue = flightPage.GetFromValue();
                Assert.IsTrue(!previousFrom.Equals(newFromValue), "From n'a pas changé");
            }
            finally
            {
                flightPage.SetFrom(previousFrom);

                if (previousValidStatus != flightPage.GetFirstFlightStatus("V"))
                {
                    flightPage.SetNewState("V");
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_HedomadaireViewSearch()
        {
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            hedomadairePage.ResetFilter();
            var firstFlightNumber = hedomadairePage.GetFirstFlightNumber().ToLower();
            hedomadairePage.Filter(HedomadairePage.FilterType.Search, firstFlightNumber);
            var totalNumberAfterSearch = hedomadairePage.CheckTotalNumber();
            var searchedFlightNumber = hedomadairePage.GetFirstFlightNumber().ToLower();
            Assert.AreEqual(searchedFlightNumber, firstFlightNumber);
            Assert.AreEqual(totalNumberAfterSearch, 1, "Le filtre ne fonctionne pas correctement.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_HedomadaireViewResetFilter()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            var firstFLightNumber = hedomadairePage.GetFirstFlightNumber();
            hedomadairePage.Filter(HedomadairePage.FilterType.Search, firstFLightNumber);
            var totalNumberBeforeReset = hedomadairePage.CheckTotalNumber();
            hedomadairePage.ResetFilter();
            var totalNumberAfterReset = hedomadairePage.CheckTotalNumber();
            Assert.IsTrue(totalNumberBeforeReset < totalNumberAfterReset, "Le filtre ne fonctionne pas correctement.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_IndexTo()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            var flightNo = flightPage.GetFirstFlightNumber();
            flightPage.Filter(FilterType.SearchFlight, flightNo);
            var previousValidStatus = flightPage.GetFirstFlightStatus("V");
            var toValue = "BCN";

            var previousTo = flightPage.GetToValue();
            try
            {
                if (previousValidStatus)
                {
                    flightPage.SetNewState("V");
                }
                if (previousTo.Equals(toValue))
                {
                    toValue = "MAD";
                }
                flightPage.SetTo(toValue);
                flightPage.Refrech();
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.SearchFlight, flightNo);
                var newToValue = flightPage.GetToValue();
                Assert.IsTrue(!previousTo.Equals(newToValue), "TO n'a pas changé");
            }
            finally
            {
                flightPage.SetTo(previousTo);

                if (previousValidStatus != flightPage.GetFirstFlightStatus("V"))
                {
                    flightPage.SetNewState("V");
                }
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_HedomadaireViewModificationQtities()
        {
            string randomGuestTypeQty = new Random().Next(1, 300).ToString();

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            hedomadairePage.ResetFilter();
            hedomadairePage.UnfoldAll();
            hedomadairePage.SetGuestTypeQty(randomGuestTypeQty);
            WebDriver.Navigate().Refresh();
            hedomadairePage.UnfoldAll();
            var NewGuestTypeQty = hedomadairePage.GetGuestTypeQty();
            var equalQty = NewGuestTypeQty.Equals(randomGuestTypeQty);
            Assert.IsTrue(equalQty, MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_HedomadaireViewStartDate()
        {
            DateTime dateInput = DateUtils.Now.AddDays(-3);

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            hedomadairePage.ResetFilter();
            hedomadairePage.Filter(HedomadairePage.FilterType.StartDate, dateInput);
            var expectedDateFormat = hedomadairePage.GetStartDateColon();
            string actualDateFormat = hedomadairePage.ConvertDateFormat(dateInput.Date).Trim();
            // Assert
            Assert.AreEqual(expectedDateFormat, actualDateFormat, "La conversion de la date n'est pas correcte.");
            var isSortedByDate = hedomadairePage.IsSortedByDate();
            Assert.IsTrue(isSortedByDate, MessageErreur.FILTRE_ERRONE, "Sort by start date");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_HedomadaireViewExport()
        {

            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersion = true;

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            hedomadairePage.ResetFilter();
            DeleteAllFileDownload();
            hedomadairePage.ClearDownloads();
            string flightNumber = hedomadairePage.GetFirstFlightNumber();
            hedomadairePage.Filter(HedomadairePage.FilterType.Search, flightNumber);
            hedomadairePage.ExportHedomadaire(HedomadairePage.ExportType.Export, newVersion);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = hedomadairePage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Flights", filePath);
            var listResult = OpenXmlExcel.GetValuesInList("Flight No", "Flights", filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsFalse(!listResult.Contains(flightNumber), MessageErreur.EXCEL_DONNEES_KO);


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_ExportGuestAndService()
        {
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["ServiceBob"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string flightNumber = new Random().Next().ToString();
            bool newVersion = true;
            string guest = "FC";

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            DeleteAllFileDownload();
            flightPage.ClearDownloads();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
            try
            {
                //Edit the first flight
                var editPage = flightPage.EditFirstFlight(flightNumber);
                editPage.AddGuestType(guest);
                editPage.AddService(serviceName);
                flightPage = editPage.CloseViewDetails();
                flightPage.ResetFilter();
                flightPage.ClearDownloads();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.ExportFlights(ExportType.ExportGuestAndServices, newVersion);
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();
                var correctDownloadedFile = flightPage.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile);
                var fileName = correctDownloadedFile.Name;
                var filePath = System.IO.Path.Combine(downloadsPath, fileName);
                int resultNumber = OpenXmlExcel.GetExportResultNumber("Guest&Services", filePath);
                //Assert
                Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_HedomadaireViewPrintServicesQuantities()
        {
            string PdfName = "PrintFlightPeriod";

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            hedomadairePage.ShowExtendedMenu();
            hedomadairePage.ClickPrintServicesQuantitiesByCustomerAndCategories();
            var titleOngletPrint = hedomadairePage.GetTitleOngletPrint();
            var isCorrect = PdfName.Equals(titleOngletPrint);
            Assert.IsTrue(isCorrect, " le fichier ne s’ouvre pas dans une nouvelle fenêtre");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_HedomadaireViewEndDate()
        {
            // ne pas avoir dateTo<dateFrom
            DateTime fixedDate = DateTime.Today;
            
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            hedomadairePage.ResetFilter();
            hedomadairePage.Filter(HedomadairePage.FilterType.EndDate, fixedDate);
            var expectedDateFormat = hedomadairePage.GetEndDateColumn();
            string actualDateFormat = hedomadairePage.ConvertDateFormat(fixedDate.Date);
            // Assert
            Assert.AreEqual(expectedDateFormat.Trim(), actualDateFormat.Trim(), "La conversion de la date n'est pas correcte.");
            Assert.IsTrue(hedomadairePage.IsSortedByDate(), MessageErreur.FILTRE_ERRONE, "Sort by start date");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_ExportFlightLogs()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ExportFlights(ExportType.ExportLogs, true);
            var exists = flightPage.CheckDownloadsExists();
            Assert.IsTrue(exists, "Le fichier n'arrive pas dans les téléchargement winrest");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_HedomadaireViewAirlinesFilter()
        {
            string airliens = "CAT Genérico";

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            hedomadairePage.ResetFilter();
            hedomadairePage.Filter(HedomadairePage.FilterType.Airlines, airliens);
            Assert.IsTrue(hedomadairePage.GetAirlinesFiltred().All(airlines => airlines == airliens), "l'application de filtre par airlines ne fonctionne pas correctement");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Export_flights_list_for_a_period()
        {
            var startDate = DateTime.Now.AddMonths(-1);
            var endDate = DateTime.Now;

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            hedomadairePage.ResetFilter();
            hedomadairePage.Filter(HedomadairePage.FilterType.StartDate, startDate);
            hedomadairePage.Filter(HedomadairePage.FilterType.EndDate, endDate);
            var flightCounter = hedomadairePage.CheckTotalNumber();
            if (flightCounter > 90)
            {
                hedomadairePage.ExportHedomadaire(HedomadairePage.ExportType.Export, true);
                var exists = hedomadairePage.CheckDownloadsExists();
                Assert.IsTrue(exists, "L'exportation a échoué : la liste d'impression est vide");

            }
            hedomadairePage.ClearDownloads();
        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_FLIG_Flight_detail_PrevaliateAll()
        {
            // prepare 
            string prefix = "FlightForPrevaliateAll";
            string flightNumber1 = prefix + new Random().Next().ToString();
            string flightNumber2 = prefix + DateTime.Now.ToString();
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);
            string etaHours = "05";
            string etdHours = "23";
            string aircraft = TestContext.Properties["Registration"].ToString();
            string prevalidateForamt = "P";
            // arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            //act
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber1, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                flightPage.WaitPageLoading();
                flightPage.UnSetNewState(prevalidateForamt);
                flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                flightPage.WaitPageLoading();
                flightPage.UnSetNewState(prevalidateForamt);
                flightPage.Filter(FilterType.SearchFlight, prefix);
                var flightDetail = flightPage.GoToFlightDetailsPage();
                flightDetail.ClickOnPrevalidateAll();
                bool check = flightDetail.CheckPrivalidate(prevalidateForamt);
                Assert.IsTrue(check,"prevalidateAll ne s'applique pas");

            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(prefix);
                massiveDeleteFlightPage.SelectSiteToFilter(site);
                massiveDeleteFlightPage.SelectCustomer(customer);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.ClickSelectAllButton();
                massiveDeleteFlightPage.Delete();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Flight_detail_ValidateAll()
        {
            // prepare 
            string prefix = "FlightForValiateAll";
            string flightNumber1 = prefix + new Random().Next().ToString();
            string flightNumber2 = prefix + DateTime.Now.ToString();
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);
            string etaHours = "05";
            string etdHours = "23";
            string aircraft = TestContext.Properties["Registration"].ToString();
            string validateFormat = "V";
            // arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            //act
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber1, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                flightPage.WaitPageLoading();
                flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                flightPage.WaitPageLoading();
                flightPage.Filter(FilterType.SearchFlight, prefix);
                var flightDetail = flightPage.GoToFlightDetailsPage();
                flightDetail.ClickOnValidateAll();
                bool check = flightDetail.Checkvalidate(validateFormat);
                Assert.IsTrue(check, "validateAll ne s'applique pas");

            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(prefix);
                massiveDeleteFlightPage.SelectSiteToFilter(site);
                massiveDeleteFlightPage.SelectCustomer(customer);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.ClickSelectAllButton();
                massiveDeleteFlightPage.Delete();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Flight_detail_InvoiceAll()
        {
            // prepare 
            string prefix = "FlightForInvoiceAll";
            string flightNumber1 = prefix + new Random().Next().ToString();
            string flightNumber2 = prefix + DateTime.Now.ToString();
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);
            string etaHours = "05";
            string etdHours = "23";
            string aircraft = TestContext.Properties["Registration"].ToString();
            string validateFormat = "V";
            string invoiceFormat = "I";
            // arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            //act
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber1, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                flightPage.WaitPageLoading();
                flightPage.SetNewState(validateFormat);
                flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                flightPage.WaitPageLoading();
                flightPage.SetNewState(validateFormat);
                flightPage.WaitPageLoading();
                flightPage.Filter(FilterType.SearchFlight, prefix);
                var flightDetail = flightPage.GoToFlightDetailsPage();
                flightDetail.ClickOnInvoieAll();
                bool check = flightDetail.CheckInvoice(invoiceFormat);
                Assert.IsTrue(check, "invoice all ne s'applique pas");

            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(prefix);
                massiveDeleteFlightPage.SelectSiteToFilter(site);
                massiveDeleteFlightPage.SelectCustomer(customer);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.ClickSelectAllButton();
                massiveDeleteFlightPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Printflightlabel()
        {
            string flightNumber = new Random().Next().ToString();
            string customer = TestContext.Properties["CustomerLpCart2"].ToString();
            string airCraft = TestContext.Properties["Aircraft"].ToString();
            string siteFrom = TestContext.Properties["SiteLpCart2"].ToString();
            string siteTo = TestContext.Properties["SiteLpCart"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.ClearDownloads();
            flightPage.Filter(FilterType.Sites, siteFrom);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, airCraft, siteFrom, siteTo);
            FlightDetailsPage FlightDetailsPage = flightPage.EditFirstFlight(flightNumber);
            FlightDetailsPage.SelectLpCart("Calomato");
            FlightDetailsPage.ShowInfoFlightExtendedMenu();
            PrintReportPage PrintReportPage = FlightDetailsPage.PrintFlightLabels();
            var windowIsOpened = PrintReportPage.VerifyIfNewWindowIsOpened();
            Assert.IsTrue(windowIsOpened, "le print n'a pas été effectuée .");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Add_Multiple_Flights_loadLP()
        {
            // Préparation des données
            string site = TestContext.Properties["SiteLpCart"].ToString();
            DateTime startDate = DateTime.Now.AddDays(-30);
            string CustomerLpFilter = TestContext.Properties["CustomerLpFilter"].ToString();
            string Aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            loadingPlanPage.ResetFilter();

            //filtre par site et par date de début
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.StartDate, startDate);

            // Vérification du nombre total de plans de chargement pour s'assurer qu'il est supérieur à 70
            var listNumber = loadingPlanPage.CheckTotalNumber();
            Assert.IsTrue(listNumber > 70, "Le nombre de loading plans n'est pas supérieur à 70.");


            FlightPage flightPage = loadingPlanPage.GoToFlights_FlightPage();
            FlightMultiCreateModalPage flightMultiplModalpage = flightPage.CreateMultiFlights();
            flightMultiplModalpage.FillField_AddMultipleFlights(site, CustomerLpFilter);

            flightMultiplModalpage.AircraftMultipleFlights(Aircraft);

            // Vérification que la liste des plans de chargement associés est chargée sans erreur ou délai d'attente
            bool isLoadingPlanLoaded = flightMultiplModalpage.IsLoadingPlanListLoaded();
            Assert.IsTrue(isLoadingPlanLoaded, "La liste de loading plan n'est pas chargée correctement ou a expiré.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Add_Multiple_Flights_goToTop()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            FlightMultiCreateModalPage flightMultiCreateModalPage = flightPage.CreateMultiFlights();
            bool result = flightMultiCreateModalPage.VerifyGoTopBtn();

            //Assert
            Assert.IsTrue(result, "Le défilement automatique au début de la popup n'a pas eu lieu.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Add_Multiple_Flights_goTobottom()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            FlightMultiCreateModalPage flightMultiCreateModalPage = flightPage.CreateMultiFlights();
            bool result = flightMultiCreateModalPage.VerifyGoToBottomBtn();

            //Assert
            Assert.IsTrue(result, "Le défilement automatique à la fin de la popup n'a pas eu lieu.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Print_popup_Periode_shifte()
        {
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.OpenDownloadPanel();
            bool isValidPanelPosition = flightPage.CheckPanelPosition();
            Assert.IsTrue(isValidPanelPosition, "The popup panel is not positioned correctly.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Flight_detail_Filter_invoiced()
        {
            // prepare
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            var flightNumber = "flightToDay" + new Random().Next();
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);
            string etaHours = "05";
            string etdHours = "23";
            string aircraft = TestContext.Properties["Registration"].ToString();
            int verif = 0;
            string invoice = "Invoice";
            string valid = "Valid";
            string validFormat = "V";
            // arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            //act
            try
            {
                // create a flight 
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                flightPage.SetNewState(validFormat);
                FlightDetailsPage flightDetail = flightPage.GoToFlightDetailsPage();
                flightDetail.FillFilterParamaters(flightNumber, site, invoice);
                flightDetail.WaitPageLoading();
                var count = flightDetail.GetNumberOfFlight();
                Assert.AreEqual(count, verif, "le filtre sur invoice ne s'applique pas");
                flightDetail.FillFilterParamaters(flightNumber, site, valid);
                flightDetail.ClickOnInvoieAll();
                flightDetail.WaitPageLoading();
                flightDetail.FillFilterParamaters(flightNumber, site, invoice);
                verif = 1;
                count = flightDetail.GetNumberOfFlight();
                Assert.AreEqual(count, verif, "le filtre sur invoice ne s'applique pas");
            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.SelectSiteToFilter(site);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Flight_detail_Filter_Preval()
        {
            // prepare
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            var flightNumber = "flightToDay" + new Random().Next();
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);
            string etaHours = "05";
            string etdHours = "23";
            string aircraft = TestContext.Properties["Registration"].ToString();
            string preval = "Preval";
            int verif = 0;
            string validFormat = "V";
            // arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            //act
            try
            {
                // create a flight 
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                flightPage.SetNewState(validFormat);
                FlightDetailsPage flightDetail = flightPage.GoToFlightDetailsPage();
                flightDetail.FillFilterParamaters(flightNumber, site, preval);
                flightDetail.WaitPageLoading();
                var count = flightDetail.GetNumberOfFlight();
                Assert.AreEqual(count, verif, "le filtre sur preval ne s'applique pas");
                flightDetail.CloseViewDetails();
                flightPage.Filter(FilterType.SearchFlight, flightNumber);
                flightPage.WaitPageLoading();
                flightPage.UnSetNewState(validFormat);
                flightDetail = flightPage.GoToFlightDetailsPage();
                flightDetail.FillFilterParamaters(flightNumber, site, preval);
                flightDetail.WaitPageLoading();
                verif = 1;
                count = flightDetail.GetNumberOfFlight();
                Assert.AreEqual(count, verif, "le filtre sur preval ne s'applique pas");

            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.SelectSiteToFilter(site);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Flight_detail_Filter_Valid()
        {
            // prepare
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            var flightNumber = "flightToDay" + new Random().Next();
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);
            string etaHours = "05";
            string etdHours = "23";
            string aircraft = TestContext.Properties["Registration"].ToString();
            string valid = "Valid";
            string validFormat = "V";
            int verif = 0;
            //arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            //act
            try
            {
                // create a flight 
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                var flightDetail = flightPage.GoToFlightDetailsPage();
                flightDetail.FillFilterParamaters(flightNumber, site, valid);
                flightDetail.WaitPageLoading();
                var count = flightDetail.GetNumberOfFlight();
                Assert.AreEqual(count,verif, "le filtre sur valid ne s'applique pas");
                flightDetail.CloseViewDetails();
                flightPage.Filter(FilterType.SearchFlight, flightNumber);
                flightPage.WaitPageLoading();
                flightPage.SetNewState(validFormat);
                flightDetail = flightPage.GoToFlightDetailsPage();
                flightDetail.FillFilterParamaters(flightNumber, site, valid);
                flightDetail.WaitPageLoading();
                verif = 1;
                count = flightDetail.GetNumberOfFlight();
                Assert.AreEqual(count,verif, "le filtre sur valid ne s'applique pas");
            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.SelectSiteToFilter(site);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Flight_detail_Filter_Cancelled()
        {
            // prepare 
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            var flightNumber = "flightToDay" + new Random().Next();
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);
            string etaHours = "05";
            string etdHours = "23";
            string aircraft = TestContext.Properties["Registration"].ToString();
            string canceled = "Cancelled";
            string canceledFormat = "C";
            int verif = 0;
            //arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            //act
            try
            {
                // create a flight 
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                var flightDetail = flightPage.GoToFlightDetailsPage();
                flightDetail.FillFilterParamaters(flightNumber, site, canceled);
                flightDetail.WaitPageLoading();
                var count = flightDetail.GetNumberOfFlight();
                Assert.AreEqual(count, verif, "le filtre sur cancelled ne s'applique pas");
                flightDetail.CloseViewDetails();
                flightPage.Filter(FilterType.SearchFlight,flightNumber);
                flightPage.WaitPageLoading();
                flightPage.SetNewState(canceledFormat);
                flightDetail = flightPage.GoToFlightDetailsPage();
                flightDetail.FillFilterParamaters(flightNumber, site, canceled);
                flightDetail.WaitPageLoading();
                verif = 1;
                count = flightDetail.GetNumberOfFlight();
                Assert.AreEqual(count, verif, "le filtre sur cancelled ne s'applique pas");
            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.SelectSiteToFilter(site);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_removeAddedService()
        {
            string siteFrom = TestContext.Properties["Site"].ToString();
            string flightNumber = new Random().Next().ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            //Edit the first flight
            var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.AddGuestType();
            var beforeResult = editPage.CountServices();
            if (beforeResult == 0)
            {
                editPage.AddGenericService();
            }
            editPage.DeleteService();
            editPage.DeleteGuestType();
            var afterResult = editPage.CountServices();
            //Assert
            if (beforeResult == 0)
            {
                Assert.AreEqual(0, afterResult, MessageErreur.SERVICE_NON_AJOUTE);
            }
            else
            {
                Assert.AreNotEqual(beforeResult, afterResult, MessageErreur.SERVICE_NON_AJOUTE);
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Add_Multiple_Flights_pickLP()
        {
            // Prepare
            string site = TestContext.Properties["SiteLpCart"].ToString();
            DateTime startDate = DateTime.Now.AddDays(-5);
            string Aircraft = TestContext.Properties["Aircraft"].ToString();
            string CustomerLpFilter = TestContext.Properties["CustomerLpFilter"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.StartDate, startDate);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Customers, CustomerLpFilter);
            var listNumber = loadingPlanPage.CheckTotalNumber();
            Assert.IsTrue(listNumber > 0, "There is no Loading Plans .");

            FlightPage flightPage = loadingPlanPage.GoToFlights_FlightPage();
            FlightMultiCreateModalPage flightMultiplModalpage = flightPage.CreateMultiFlights();
            flightMultiplModalpage.FillField_AddMultipleFlights(site, CustomerLpFilter);
            flightMultiplModalpage.AircraftMultipleFlights(Aircraft);
            flightMultiplModalpage.pickMultipleLodingPlan();
            var legDetails = flightMultiplModalpage.GetFirstLegDetails();
            Assert.IsTrue(legDetails != "No leg associated", "Leg details were not updated correctly.");

            var guestDetails = flightMultiplModalpage.GetFirstGuestDetails();
            Assert.IsTrue(!string.IsNullOrEmpty(guestDetails), "Guest details are missing.");

            string qtyValue = flightMultiplModalpage.GetQtyDetails();
            Assert.IsTrue(!string.IsNullOrEmpty(qtyValue), "Quantity details are missing.");
            Assert.IsTrue(int.TryParse(qtyValue, out int quantity) && quantity > 0, "Quantity value is invalid.");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Customer_Price_serviceIs_0()
        {
            // Prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customerBob = TestContext.Properties["Bob_CustomerFilter"].ToString();
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);
            bool newVersion = true;
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string qtyForService = "0";
            string price = "0";

            HomePage homePage = LogInAsAdmin();

            try
            {
                ServicePage servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.BackToList();
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerBob, fromDate, toDate);
                pricePage.ToggleFirstPrice();
                ServiceCreatePriceModalPage serviceEditPriceModal = servicePricePage.ServiceEitPriceModal(1);
                priceModalPage.SetMethodScaleWithZeroNbPersonsWithoutSAVE("Scale", 1, price);
                servicePricePage = priceModalPage.Save();
                servicePricePage.BackToList();

                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);

                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);

                var flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
                flightDetailsPage.AddGuestTypeFlight();
                flightDetailsPage.AddService(serviceName);
                flightDetailsPage.SetFinalQtyServiceFlight(qtyForService);
                WebDriver.Navigate().Refresh();
                DeleteAllFileDownload();
                flightPage.ClearDownloads();
                flightPage.ExportFlights(ExportType.Export, newVersion);

                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();
                var correctDownloadedFile = flightPage.GetExportExcelFile(taskFiles);

                Assert.IsNotNull(correctDownloadedFile);
                var fileName = correctDownloadedFile.Name;
                var filePath = System.IO.Path.Combine(downloadsPath, fileName);

                int resultNumber = OpenXmlExcel.GetExportResultNumber("Flights", filePath);
                var listResult = OpenXmlExcel.GetValuesInList("Flight No", "Flights", filePath);
                var listResultService = OpenXmlExcel.GetValuesInList("Service", "Flights", filePath);

                //Assert
                Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
                Assert.IsFalse(!listResult.Contains(flightNumber), MessageErreur.EXCEL_DONNEES_KO);
                Assert.IsFalse(!listResultService.Contains(serviceName), MessageErreur.EXCEL_DONNEES_KO);
            }
            finally
            {
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, null, null);
                WebDriver.Navigate().Refresh();
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.ResetFilters();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Flights_Apply_LPCart_to_future_same_flights()
        {
            //constants LpCart
            Random rnd = new Random();
            var nameLpCart = "lpCart_" + DateUtils.Now.ToString("dd/MM/yyyy");
            string siteACE = TestContext.Properties["SiteLpCart"].ToString();
            string customerFilter = TestContext.Properties["CustomerLPFlight"].ToString();
            string customerCAT = "$$ - CAT Genérico";
            string airCraftB777 = TestContext.Properties["Aircraft"].ToString();
            string code = DateTime.Now.ToString("yyyy-MM-dd") + "-" + rnd.Next(0, 500).ToString();
            DateTime from = DateUtils.Now.AddMonths(-2);
            DateTime to = DateUtils.Now.AddMonths(2);
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            string routes = TestContext.Properties["RouteLP"].ToString();
            //constants Flight
            string flightNumber = new Random().Next().ToString();
            string siteToBCN = TestContext.Properties["SiteToFlight"].ToString();

            HomePage homePage = LogInAsAdmin();
            //Act

            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            LpCartCreateModalPage lpCartCreateModal = lpCartPage.LpCartCreatePage();
            LpCartCartDetailPage lpCartCartDetailPage = lpCartCreateModal.FillField_CreateNewLpCartWithRoutes(code, nameLpCart, siteACE, customerCAT, airCraftB777, from, to, comment, routes);
            lpCartCartDetailPage.BackToList();
            homePage.Navigate();

            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            try
            {
                FlightMultiCreateModalPage flightMultiplModalpage = flightPage.CreateMultiFlights();
                flightMultiplModalpage.FillField_CreatMultiFlight(siteACE, flightNumber, siteToBCN, airCraftB777, DateUtils.Now, DateUtils.Now.AddDays(30));
                flightMultiplModalpage.SetETA("16:00", 1);
                flightMultiplModalpage.SetETD("13:00", 1);
                flightPage = flightMultiplModalpage.Validate();
                flightPage.Filter(FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FilterType.Sites, siteACE);
                FlightDetailsPage flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
                flightDetailsPage.SelectLpCart(nameLpCart);
                flightDetailsPage.ApplyLPCartToSameFutureFlightsWithoutConfirm(true);
                bool isModalOpen = flightDetailsPage.ModalLPCartToSameFutureFlightsInfo();
                flightDetailsPage.ConfirmApplyLPCartToSameFutureFlights();
                Assert.IsTrue(isModalOpen, "LPCart is not applied to the same future flights.");
                flightPage = flightDetailsPage.CloseViewDetails();
            }
            finally
            {
                flightPage  = flightPage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.MassiveDeleteMenus(flightNumber, siteACE, customerFilter, true);
                lpCartPage = flightPage.GO_To_LPCART_Tab();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, nameLpCart);
                if (lpCartPage.CheckTotalNumber() >= 1)
                {
                    lpCartPage.DeleteLpCart();
                }
            }

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_addRepeatedService()
        {

            string preval = "Preval";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            WebDriver.Manage().Window.Maximize();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.ShowStatusMenu();
            flightPage.Filter(FlightPage.FilterType.Status, preval);
            string resultPreval = flightPage.GetFirstFlightNumber();
            var editPage = flightPage.EditFirstFlight(resultPreval);

            if (editPage.CountServices() != 0)
            {
                string service = editPage.GetServiceName();
                editPage.AddService(service);

            }
            else
            {
                editPage.AddGuestTypeFlight();
                editPage.AddService("AIR FRESHNER TVS");
                editPage.AddService("AIR FRESHNER TVS");

            }
            bool isServiceExist = editPage.CheckServiceExists();
            Assert.IsTrue(isServiceExist, "Le message d'erreur 'service already exists' n'apparaît pas ");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_duplicated_Legs_corresponding_tothedates_of_thenewFlights()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string flightNumber = new Random().Next().ToString();
            DateTime newFlightDateFrom = DateTime.Now.AddDays(10);
            DateTime newFlightDateTo = newFlightDateFrom.AddDays(7);
            var downloadPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.ClearDownloads();
            flightPage.PurgeDownloads();
            try
            {
                //Create Multi Flight
                var flightMultiplModalpage = flightPage.CreateMultiFlights();
                flightMultiplModalpage.FillField_CreatMultiFlight(siteFrom, flightNumber, siteTo, aircraft);
                flightPage = flightMultiplModalpage.Validate();

                //Apply a filter on created flight
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var flightEditPage = flightPage.EditFlightRow(flightNumber);
                flightEditPage.ClickOnAddGuestType();
                flightEditPage.AddNewService();
                flightEditPage.DuplicateFlight(flightNumber, newFlightDateFrom, newFlightDateTo);
                // export processing for flight one 
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.SetDateState(DateTime.Now);
                flightPage.ClearDownloads();
                flightPage.PurgeDownloads();
                flightPage.Export(ExportType.Export);
                var correctDownloadedFile = flightPage.GetExportExcelFile(new DirectoryInfo(downloadPath).GetFiles());
                Assert.IsNotNull(correctDownloadedFile);
                var fileName = correctDownloadedFile.Name;
                var filePath = System.IO.Path.Combine(downloadPath, fileName);
                var flightLegs = OpenXmlExcel.GetValuesInList("Leg Name", "Flights", correctDownloadedFile.FullName);

                // export processing for flight two 
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.SetDateState(newFlightDateFrom);
                flightPage.ClearDownloads();
                flightPage.PurgeDownloads();
                flightPage.Export(ExportType.Export);
                correctDownloadedFile = flightPage.GetExportExcelFile(new DirectoryInfo(downloadPath).GetFiles());
                Assert.IsNotNull(correctDownloadedFile);
                fileName = correctDownloadedFile.Name;
                filePath = System.IO.Path.Combine(downloadPath, fileName);
                var dupflightLegs = OpenXmlExcel.GetValuesInList("Leg Name", "Flights", correctDownloadedFile.FullName);

                //Assert
                Assert.AreEqual(flightLegs[0], dupflightLegs[0], "les nouveaux Legs correspondants aux dates des vols dupliqués n'ont pas les memes");
            }
            finally
            {
                flightPage.ResetFilter();
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_ChangeStatus_Prevalide()
        {
            //Prepare
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customer = TestContext.Properties["Bob_Customer"].ToString();
            string flightNumber = new Random().Next().ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, site);

            try
            {
                //create flight
                flightPage.ShowPlusMenu();
                FlightCreateModalPage flightCreateModalPage = flightPage.FlightCreatePage();
                flightCreateModalPage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FilterType.Sites, site);

                var flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
                flightDetailsPage.AddGuestType();
                flightDetailsPage.AddService(serviceName);

                //edit status
                flightDetailsPage.ClickOnPrevalidate();
                flightDetailsPage.CloseViewDetails();

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FilterType.Sites, site);
                bool isPervalClicked = flightPage.CheckChangedStatus("P");

                //Assert
                Assert.IsFalse(isPervalClicked, "Le status n'a pas changé");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.Sites, site);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, site, customer,true);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_ChangeStatus_Validate()
        {
            //Prepare
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customer = TestContext.Properties["Bob_Customer"].ToString();
            string flightNumber = new Random().Next().ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, site);

            try
            {
                //create flight
                flightPage.ShowPlusMenu();
                FlightCreateModalPage flightCreateModalPage = flightPage.FlightCreatePage();
                flightCreateModalPage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FilterType.Sites, site);
                bool isPervalClicked = flightPage.CheckChangedStatus("P");
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);

                var flightDetailsPage = flightPage.EditFlightByName(flightNumber);
                flightDetailsPage.AddGuestType();
                flightDetailsPage.AddService(serviceName);
                if (!isPervalClicked)
                {
                    flightDetailsPage.ClickOnPrevalidate();

                }
                flightDetailsPage.CloseViewDetails();
                flightDetailsPage = flightPage.EditFlightByName(flightNumber);
                flightDetailsPage.SetNewState("V");
                flightDetailsPage.CloseViewDetails();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FilterType.Sites, site);
                bool isValidateClicked = flightPage.CheckChangedStatus("V");

                //Assert
                Assert.IsTrue(isValidateClicked, "Le status n'a pas changé");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.Sites, site);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, site, customer, true);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_ChangeStatus_Invoice()
        {
            //Prepare
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customer = TestContext.Properties["Bob_Customer"].ToString();
            string flightNumber = new Random().Next().ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, site);

            try
            {
                //create flight
                flightPage.ShowPlusMenu();
                FlightCreateModalPage flightCreateModalPage = flightPage.FlightCreatePage();
                flightCreateModalPage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FilterType.Sites, site);
                bool isPervalClicked = flightPage.CheckChangedStatus("P");
                bool isValidateClicked = flightPage.CheckChangedStatus("V");
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);

                var flightDetailsPage = flightPage.EditFlightByName(flightNumber);
                flightDetailsPage.AddGuestType();
                flightDetailsPage.AddService(serviceName);
                if (!isPervalClicked)
                {
                    flightDetailsPage.SetNewState("P");
                }
                flightDetailsPage.CloseViewDetails();
                flightDetailsPage = flightPage.EditFlightByName(flightNumber);
                if (!isValidateClicked)
                {
                    flightDetailsPage.SetNewState("V");
                    flightDetailsPage.CloseViewDetails();
                    flightDetailsPage = flightPage.EditFlightByName(flightNumber);
                }

                flightDetailsPage.SetNewState("I");
                flightDetailsPage.CloseViewDetails();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FilterType.Sites, site);

                bool isInvoiceClicked = flightPage.CheckChangedStatus("I");

                //Assert
                Assert.IsTrue(isInvoiceClicked, "Le status n'a pas changé");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.Sites, site);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, site, customer, true);
            }
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_ChangeStatus_Cancel()
        {
            //Prepare
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customer = TestContext.Properties["Bob_Customer"].ToString();
            string flightNumber = new Random().Next().ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Sites, site);

            try
            {
                //create flight
                flightPage.ShowPlusMenu();
                FlightCreateModalPage flightCreateModalPage = flightPage.FlightCreatePage();
                flightCreateModalPage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FilterType.Sites, site);

                bool isPervalClicked = flightPage.CheckChangedStatus("P");

                var flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
                flightDetailsPage.AddGuestType();
                flightDetailsPage.AddService(serviceName);
                flightPage = flightDetailsPage.CloseViewDetails();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FilterType.Sites, site);
                flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
                if (!isPervalClicked)
                {
                    flightDetailsPage.SetNewState("P");  
                    flightDetailsPage.CloseViewDetails();
                    flightPage.ResetFilter();
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                    flightPage.Filter(FilterType.Sites, site);
                    flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
                }
                
                flightDetailsPage.SetNewState("V");
                flightDetailsPage.CloseViewDetails();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FilterType.Sites, site);
                flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
                flightDetailsPage.SetNewState("I");
                flightDetailsPage.ClickCancelledFlight();
                flightDetailsPage.WaitLoading();
                bool checkCancelled  = flightDetailsPage.CheckCancelled("C");

                flightDetailsPage.CloseViewDetails();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FilterType.Sites, site);

                //Assert
                flightPage.ShowStatusCancelled();
                isPervalClicked = flightPage.CheckChangedStatus("P");
                bool isValidateClicked = flightPage.CheckChangedStatus("V");
                bool isInvoiceClicked = flightPage.CheckChangedStatus("I");
               
                Assert.IsFalse(isInvoiceClicked || isValidateClicked || isPervalClicked, "Le status n'est pas effectué correctement");
                Assert.IsTrue(checkCancelled, "Le Cancelled n'est pas effectué correctement");

            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.Sites, site);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, site, customer, true);
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_SuppressionService_pop_up_flight_Out()
        {
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string qtyGuest = "10";

            var homePage = LogInAsAdmin();
            // Act Flight
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var editPage = flightPage.EditFirstFlight(flightNumber);
                // Edit flight
                editPage.AddGuestType("BOB");
                editPage.AddService(null);
                editPage.SetFinalQty(qtyGuest);
                // Suppression du service
                editPage.DeleteService();
                bool IsPopupOpen = editPage.IsFlightDetailsPopupOpen();
                // Assert que la pop-up de détail du vol est toujours ouverte
                Assert.IsTrue(IsPopupOpen, "La fenêtre de détail du vol est fermée après la suppression du service.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumber);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_NumberAlphabetic()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumberAndLetters = "FL" + new Random().Next(1000, 9999).ToString(); // e.g., FL1234

            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalPage = flightPage.FlightCreatePage();
                flightCreateModalPage.FillField_CreatNewFlight(flightNumberAndLetters, customer, aircraft, siteFrom, siteTo);
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberAndLetters);
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var displayedNumber = flightPage.GetFirstFlightNumber();
                Assert.AreEqual(flightNumberAndLetters, displayedNumber, "Le système n'a pas correctement géré le numéro de vol avec des lettres et des chiffres.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNumberAndLetters);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_UseCase_finalFigures()
        {
            // Prepare
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteAalTrolley"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string AircraftLpCart = TestContext.Properties["AircraftLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string customerCode = TestContext.Properties["CustomerScheduleCode"].ToString();
            string flightNumber = new Random().Next().ToString();
            string preval = "Preval";


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);


            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo, null, "04", "22");
            FlightDetailsPage flightEditPage = flightPage.EditFirstFlight(flightNumber);
            flightEditPage.AddGuestType();
            flightEditPage.AddService("AIR FRESHNER TVS");
            flightEditPage.CloseViewDetails();

            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Status, preval);

            //filtrage Par Site
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var from = flightPage.VerifySiteFrom(siteFrom);
            Assert.IsTrue(from, "erreur de filtrage dans le site");

            //filtrage Par Client 
            flightPage.Filter(FlightPage.FilterType.Customer, customer);
            Assert.IsTrue(flightPage.VerifyCustomer(customerCode), MessageErreur.FILTRE_ERRONE, "Customers");

            //filtrage Par ETD From 
            flightPage.Filter(FlightPage.FilterType.ETDFrom, "0400AM");
            Assert.IsTrue(flightPage.CheckTotalNumber() > 0, MessageErreur.FILTRE_ERRONE, "ETD From");

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.SetAC(AircraftLpCart);
            flightPage.SetAC(AircraftLpCart);
            var newAircraft = flightPage.GetACValue();

            Assert.AreEqual(newAircraft, AircraftLpCart, MessageErreur.FILTRE_ERRONE, "Aircraft");

            FlightDetailsPage flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
            flightDetailsPage.SetFinalQty("5");
            Assert.IsTrue(flightDetailsPage.GetServiceFinalQty1("5"), "The Quantity didn't change");

            flightDetailsPage.SetNewState("V");
            Assert.IsTrue(flightDetailsPage.GetFlightStatus("V"), "not validate");

            flightDetailsPage.SetNewState("I");
            Assert.IsTrue(flightDetailsPage.GetFlightStatus("I"), "not Invoice");

            flightDetailsPage.Filter(FlightDetailsPage.FilterType.SearchFlight, " ");
            flightDetailsPage.ClickFirstFlightInList(flightNumber);
            var newFlight = flightDetailsPage.GetFlightNumber();
            Assert.AreNotEqual(newFlight, flightNumber, "Le vol n'a pas été modifié.");



        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Add_Multiple_Flights_addNew()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string flightNumber = new Random().Next().ToString();
            string messageError = "A flight with the same name already exists for this date";

            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightMultiplModalpage = flightPage.CreateMultiFlights();
            flightMultiplModalpage.FillField_CreatMultiFlight(siteFrom, flightNumber, siteTo, aircraft);

            flightPage = flightMultiplModalpage.AddNew();
            flightPage = flightMultiplModalpage.AddNew();

            bool isAfficheMsgError = flightMultiplModalpage.AfficheMsgError(messageError);
            bool isFlightUpdate = flightMultiplModalpage.UpdateDateFlightVerifInformation(siteFrom, flightNumber, siteTo, DateTime.Today.AddDays(1));
            flightPage = flightMultiplModalpage.AddNew();

            bool IsAddNewFlightMemory = flightMultiplModalpage.AddNewFlightMemory();

            Assert.IsTrue(isAfficheMsgError, "Une erreur doit apparaître qui m'empêche de continuer.");
            Assert.IsTrue(IsAddNewFlightMemory, "Le système n'a pas la tendance à devenir lent");
            Assert.IsTrue(isFlightUpdate, "Les informations du vol antérieur restent sur place pour être modifiables.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_ClosingPeriod()
        {
            //Prepare
            string sectionToCheck = "Closing Periods";

            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var globalSettingPage = homePage.GoToParameters_GlobalSettings();
            bool isClosingPeriodVisible = globalSettingPage.DetectClosingPeriod(sectionToCheck);

            //Assert
            Assert.IsFalse(isClosingPeriodVisible, "Closing period existe sous Global settings");

            var flightPage = homePage.GoToFlights_FlightPage();
            bool isClosingPeriodVisibleInFlights = flightPage.DetectClosingPeriod(sectionToCheck);

            //Assert
            Assert.IsFalse(isClosingPeriodVisibleInFlights, "Closing period existe sous Flights");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void FL_FLIG_HedomadaireViewETDFrom()
        {
            string ETDFrom = "0500AM";
            HomePage homePage = LogInAsAdmin();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            hedomadairePage.ResetFilter();
            int flightCounterBefore = hedomadairePage.CheckTotalNumber();
            hedomadairePage.Filter(HedomadairePage.FilterType.ETDFrom, ETDFrom);
            int flightCounterbeAfter = hedomadairePage.CheckTotalNumber();
            Assert.AreNotEqual(flightCounterBefore, flightCounterbeAfter, MessageErreur.FILTRE_ERRONE, "l'application de filtre par ETD From ne fonctionne pas correctement.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void FL_FLIG_HedomadaireViewETDTo()
        {
            string ETDTo = "0500PM";
            HomePage homePage = LogInAsAdmin();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            hedomadairePage.ResetFilter();
            int flightCounterBefore = hedomadairePage.CheckTotalNumber();
            hedomadairePage.Filter(HedomadairePage.FilterType.ETDTo, ETDTo);
            int flightCounterbeAfter = hedomadairePage.CheckTotalNumber();
            Assert.AreNotEqual(flightCounterBefore, flightCounterbeAfter, MessageErreur.FILTRE_ERRONE, "l'application de filtre par ETD To ne fonctionne pas correctement.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_EditQtyOfGuestType()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string flightNumberBis = flightNumber + "_bis";
            string editedGuestTypeQty = "5";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            try
            {
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

                //Edit the Guesttype quantity
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var flightEditPage = flightPage.EditFirstFlight(flightNumber);
                flightEditPage.AddGuestType();
                string guestTypeQT = flightEditPage.GetGuestTypeQty();

                flightEditPage.EditGuestTypeQT(editedGuestTypeQty);

                flightPage = flightEditPage.CloseViewDetails();
                flightEditPage = flightPage.EditFirstFlight(flightNumber);
                guestTypeQT = flightEditPage.GetGuestTypeQty();

                //Assert
                Assert.AreEqual(editedGuestTypeQty, guestTypeQT, "la quantité de Guest Type n'a pas été modifiée.");

            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, siteFrom, customer, false);
                flightPage.ResetFilter();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_EditAircraft()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string flightNumberBis = flightNumber + "_bis";
            string editedAircraft = TestContext.Properties["AircraftBis"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            try
            {
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

                //Edit Flight Aircraft
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var flightEditPage = flightPage.EditFirstFlight(flightNumber);
                flightEditPage.ChangeAircaftForLoadingPlan(editedAircraft);

                flightPage = flightEditPage.CloseViewDetails();
                flightEditPage = flightPage.EditFirstFlight(flightNumber);
                string flightAircraft = flightEditPage.GetFlightAircraft();

                //Assert
                Assert.AreEqual(editedAircraft, flightAircraft, "Aircraft n'a pas été modifiée.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, siteFrom, customer, false);
                flightPage.ResetFilter();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_EditArrFlightNumber()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string editedArrFlightNumber = new Random().Next().ToString(); 

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            try
            {
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

                //Edit ArrFlightNumber
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var flightEditPage = flightPage.EditFirstFlight(flightNumber);
                flightEditPage.AddArrFlightNumber(editedArrFlightNumber);
                flightPage = flightEditPage.CloseViewDetails();

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightEditPage = flightPage.EditFirstFlight(flightNumber);

                string arrNumber = flightEditPage.GetArrFlightNumber();

                //Assert
                Assert.AreEqual(editedArrFlightNumber, arrNumber, "Arr Flight Number n'a pas été modifiée.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, siteFrom, customer, false);
                flightPage.ResetFilter();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_EditRegistration()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string editedRegistration = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            try
            {
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

                //Edit ArrFlightNumber
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var flightEditPage = flightPage.EditFirstFlight(flightNumber);
                flightEditPage.SetRegistration(editedRegistration);
                flightPage = flightEditPage.CloseViewDetails();

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightEditPage = flightPage.EditFirstFlight(flightNumber);

                string registrationValue = flightEditPage.GetRegistration();

                //Assert
                Assert.AreEqual(editedRegistration, registrationValue, "La Registration n'a pas été modifiée.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, siteFrom, customer, false);
                flightPage.ResetFilter();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_EditTabletFlightType()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string editTabletFlightType = "International";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            try
            {
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);

                //Edit TabletFlightType
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var flightEditPage = flightPage.EditFirstFlight(flightNumber);
                flightEditPage.SetTabletFlightType(editTabletFlightType);
                flightPage = flightEditPage.CloseViewDetails();

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightEditPage = flightPage.EditFirstFlight(flightNumber);

                string registrationValue = flightEditPage.GetTabletFlightType();

                //Assert
                Assert.AreEqual(editTabletFlightType, registrationValue, "Tablet Flight Type n'a pas été modifiée.");
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, siteFrom, customer, false);
                flightPage.ResetFilter();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_Detail_EditLpCart()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customerLpCart = "$$ - CAT Genérico";
            string customer = "CAT Genérico";
            string flightNumber = new Random().Next().ToString();

            var lpCartName= $"LP Cart{DateTime.Now.ToString("dd/MM/yyyy")}";
            var code = new Random()
                .Next(0, 5000)
                .ToString();
            var from = DateTime.Today.AddDays(-700);
            var to = DateTime.Today.AddDays(600);
            var route = TestContext.Properties["Route"].ToString();
            string comment = TestContext.Properties["CommentLpCart"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            bool isLPCartCreated = false;
            bool isFlightCreated = false;
            //Act
            try
            {
                //Create LpCart
                LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var lpCartCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCartWithRoutes(
                    code, lpCartName, siteFrom, customerLpCart, aircraft, from, to, comment, route);
                isLPCartCreated = true;

                // create flight
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
                isFlightCreated = true;

                //Edit LpCart
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var flightEditPage = flightPage.EditFirstFlight(flightNumber);
                flightEditPage.SelectLpCart(lpCartName);
                flightPage = flightEditPage.CloseViewDetails();

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightEditPage = flightPage.EditFirstFlight(flightNumber);

                string editedLpCart = flightEditPage.GetLpCart();

                //Assert
                Assert.AreEqual(lpCartName, editedLpCart, "LpCart n'a pas été modifiée.");
            }
            finally
            {
                if (isFlightCreated)
                {
                    var flightPage = homePage.GoToFlights_FlightPage();
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                    flightPage.MassiveDeleteMenus(flightNumber, siteFrom, customer, false);
                    flightPage.ResetFilter();
                }

                if (isLPCartCreated)
                {
                    var lpCartPage = homePage.GoToFlights_LpCartPage();
                    lpCartPage.ResetFilter();
                    lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                    lpCartPage.DeleteLpCart();
                    lpCartPage.ResetFilter();
                }
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PictoDatasheetProducedService()
        {
            //Prepare
            string guestType = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string datasheetName = "Datasheet" + "-" + DateTime.Now;// new Random().Next().ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            //string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();

            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string serviceNameNoInvoiced = TestContext.Properties["ServiceBob"].ToString();
            string serviceNameInvoiced = TestContext.Properties["ServiceNameLP"].ToString();
            string editedGuestTypeQty = "10";
            string code = new Random().Next().ToString();
            DateTime fromDate = DateTime.Today;
            DateTime toDate = DateTime.Today.AddMonths(1);

            string serviceProduction = GenerateName(4);
            string serviceName = serviceNameToday + new Random(1000).Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            try
            {
                //Create datasheet
                var datasheet = homePage.GoToMenus_Datasheet();
                var datasheetCreateModalPage = datasheet.CreateNewDatasheet();
                datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

                //Create service
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, code, serviceProduction, null, guestType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate, "10", datasheetName);

                var generalInfo = pricePage.ClickOnGeneralInformationTab();
                if (!generalInfo.IsProduced())
                {
                    generalInfo.SetProduced(true);
                }
                servicePage = generalInfo.BackToList();

                //Create flight
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);

                //Edit the Guesttype quantity
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var flightEditPage = flightPage.EditFirstFlight(flightNumber);
                flightEditPage.AddGuestType();
                flightEditPage.AddService(serviceName);

                flightEditPage.EditGuestTypeQT(editedGuestTypeQty);
                bool datasheetExists = flightEditPage.VerifyDatasheetExist();
                flightPage = flightEditPage.CloseViewDetails();

                //Assert
                Assert.IsTrue(datasheetExists, "Le picto n'est pas présent en vert.");
            }
            finally
            {
                // delete flight
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, site, customer, false);
                flightPage.ResetFilter();

                //delete datasheet
                var datasheet = homePage.GoToMenus_Datasheet();
                var massiveDelete = datasheet.OpenMassiveDeletePopup();
                massiveDelete.SetDatasheetName(datasheetName);
                massiveDelete.ClickOnSearch();
                massiveDelete.ClickOnSelectAll();
                massiveDelete.ClickDeleteButton();

                //delete service
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PictoWithoutDatasheetProducedService()
        {
            //Prepare
            string guestType = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string editedGuestTypeQty = "10";
            string code = new Random().Next().ToString();
            DateTime fromDate = DateTime.Today;
            DateTime toDate = DateTime.Today.AddMonths(1);

            string serviceProduction = GenerateName(4);
            string serviceName = serviceNameToday + new Random(2000).Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            try
            {
                //Create service
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, code, serviceProduction, null, guestType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate, "10");

                var generalInfo = pricePage.ClickOnGeneralInformationTab();
                if (!generalInfo.IsProduced())
                {
                    generalInfo.SetProduced(true);
                }
                servicePage = generalInfo.BackToList();

                //Create flight
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);

                //Edit the Guesttype quantity
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var flightEditPage = flightPage.EditFirstFlight(flightNumber);
                flightEditPage.AddGuestType();
                flightEditPage.AddService(serviceName);

                flightEditPage.EditGuestTypeQT(editedGuestTypeQty);
                bool picoExists = flightEditPage.VerifyPicoExist();
                flightPage = flightEditPage.CloseViewDetails();

                //Assert
                Assert.IsTrue(picoExists, "Le picto n'est pas présent en orangé.");
            }
            finally
            {
                // delete flight
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, site, customer, false);
                flightPage.ResetFilter();

                //delete service
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_PictoDatasheetServiceNotProduced()
        {
            //Prepare
            string guestType = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string datasheetName = "Datasheet" + "-" + DateTime.Now;// new Random().Next().ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            //string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string flightNumber = new Random().Next().ToString();

            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string serviceNameNoInvoiced = TestContext.Properties["ServiceBob"].ToString();
            string serviceNameInvoiced = TestContext.Properties["ServiceNameLP"].ToString();
            string editedGuestTypeQty = "10";
            string code = new Random().Next().ToString();
            DateTime fromDate = DateTime.Today;
            DateTime toDate = DateTime.Today.AddMonths(1);

            string serviceProduction = GenerateName(4);
            string serviceName = "Service" + "-" +DateTime.Now;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            try
            {
                //Create datasheet
                var datasheet = homePage.GoToMenus_Datasheet();
                var datasheetCreateModalPage = datasheet.CreateNewDatasheet();
                datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

                //Create service
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, code, serviceProduction, null, guestType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate, "10", datasheetName);

                var generalInfo = pricePage.ClickOnGeneralInformationTab();
                generalInfo.SetNotProduced();
                servicePage = generalInfo.BackToList();

                //Create flight
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);

                //Edit the Guesttype quantity
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                var flightEditPage = flightPage.EditFirstFlight(flightNumber);
                flightEditPage.AddGuestType();
                flightEditPage.AddService(serviceName);

                flightEditPage.EditGuestTypeQT(editedGuestTypeQty);
                bool picoExists = flightEditPage.PicoExist();
                flightPage = flightEditPage.CloseViewDetails();

                //Assert
                Assert.IsFalse(picoExists, "Le picto est pas présent.");
            }
            finally
            {
                // delete flight
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.MassiveDeleteMenus(flightNumber, site, customer, true);
                flightPage.ResetFilter();

                //delete datasheet
                var datasheet = homePage.GoToMenus_Datasheet();
                var massiveDelete = datasheet.OpenMassiveDeletePopup();
                massiveDelete.SetDatasheetName(datasheetName);
                massiveDelete.ClickOnSearch();
                massiveDelete.ClickOnSelectAll();
                massiveDelete.ClickDeleteButton();

                //delete service
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_AddMultipleFlightDates()
        {
            // Prepare 
            DateTime date = DateTime.Today;

            // Arrange
            var homePage = LogInAsAdmin();
            //Act
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            FlightMultiCreateModalPage flightMultiplModalpage = flightPage.CreateMultiFlights();
            flightMultiplModalpage.Create();

            var isVisibleYoumustdefineAtLeastOneRule = flightMultiplModalpage.isVisibleYoumustdefineAtLeastOneRule();
            //Assert
            Assert.IsTrue(isVisibleYoumustdefineAtLeastOneRule, "You must define at least one rule n'est pas visible");

            flightMultiplModalpage.ClickOnNewRule();

            string dateString = flightMultiplModalpage.GetDate();

            DateTime dateStart;
            bool isValidDate = DateTime.TryParseExact(dateString, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateStart);

            if (!isValidDate)
            {
                Assert.Fail("The date format is invalid or could not be parsed: " + dateString);
            }
            //Assert
            Assert.AreEqual(dateStart.Date, date.Date, "The flight date does not match today's date.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_ViewFlightIndex_FilterEquipmenttype_Unknown()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ShowEquipmentTypeMenu();
            flightPage.Filter(FlightPage.FilterType.EQUIPMENTTYPE_Unknown, true);
            flightPage.Filter(FlightPage.FilterType.EQUIPMENTTYPE_KSSU, false);
            flightPage.Filter(FlightPage.FilterType.EQUIPMENTTYPE_Atlas, false);
            flightPage.PageUp();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            bool isExist = flightPage.IsFlightExist();
            Assert.IsTrue(isExist,"List Flights with equipment type 'UNKOWN' empty");
            string flightNumberFromInterface = flightPage.GetFirstFlightNumber();
            Assert.AreEqual(flightNumber, flightNumberFromInterface,$"the flight {flightNumber} not exist");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_ViewFlight_FilterHidecancelledflights()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            int flightNumberBeforeFilter = flightPage.GetCountResult();
            flightPage.Filter(FlightPage.FilterType.HideCancelledFlight, true);
            int flightNumberAfterFilter = flightPage.GetCountResult();
            Assert.IsTrue(flightNumberAfterFilter < flightNumberBeforeFilter, "Le filtre n est pas appliquer");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_ViewFlightIndex_FilterCustomers()
        {
            string customer = TestContext.Properties["Customer"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.Customer, customer);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            bool isExist = flightPage.IsFlightExist();
            Assert.IsTrue(isExist, $"Flight {flightNumber} with customer {customer} not created");
            string firstFlightNumber = flightPage.GetFirstFlightNumber();
            Assert.AreEqual(firstFlightNumber, flightNumber, "Le filtre n est pas appliquer");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_ViewFlightIndex_FilterDocknumber()
        {
            string dockNumber = new Random().Next(10,99).ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.SearchFlight, flightNumber);
            flightPage.CliCkOnFirstFlight(2);
            flightPage.SetDockNumber(dockNumber);
            flightPage.ResetFilter();
            flightPage.ScrollUntilDockNumberFilterIsVisible();
            flightPage.Filter(FilterType.DockNumber, dockNumber);
            bool isExist = flightPage.IsFlightExist();
            Assert.IsTrue(isExist, $"Flight {flightNumber} with Dock Number {dockNumber} not updated");
            string isfiltred = flightPage.GetIndexDockNumberValue();
            Assert.AreEqual(isfiltred, dockNumber, $"Filter Dock Number : {dockNumber} not appliqued");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_ViewFlightIndex_FilterDriver()
        {
            //Arrange
            string driver = "CC";
            //act
            HomePage homePage = LogInAsAdmin();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FilterType.SearchFlight, flightNumber);
            flightPage.CliCkOnFirstFlight(2);
            flightPage.SetDriver(driver);
            flightPage.ResetFilter();
            //Filter
            flightPage.Filter(FilterType.Driver, driver);
            bool isExist = flightPage.IsFlightExist();
            //assert
            Assert.IsTrue(isExist, $"Le vol {flightNumber} avec le conducteur {driver} n'a pas été mis à jour.");
            string FiltredResult = flightPage.GetIndexDriverValue();
            Assert.AreEqual(FiltredResult, driver, $"Le filtre CONDUCTEUR : {driver} n'a pas été appliqué.");

        }
        [TestMethod]
        [Timeout(600000)]
        public void FL_FLIG_Calendar_FilterSite()
        {
            //Prepare
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string siteACE = "ACE";
            string serviceName = "PirlouitS";
            //login
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType();
            editPage.AddService(serviceName);
            flightPage = editPage.CloseViewDetails();
            HedomadairePage hedomadairePage = flightPage.GoToCalendarView();
            // Récupérer les sites des vols affichés
            var displayedFlights = hedomadairePage.GetDisplayedFlights();
            // Vérifier que tous les vols affichés correspondent au site
            bool displayedFlight = displayedFlights.Any(site => site.Contains(siteACE));
            Assert.IsTrue(displayedFlight, "Tous les vols affichés doivent correspondre au site sélectionné");

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIGHT_Modification_State_V()
        {
            // prepare 
            string prefix = "FlightForValiateAll";
            string flightNumber1 = prefix + new Random().Next().ToString();
            string flightNumber2 = prefix + DateTime.Now.ToString();
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);
            string etaHours = "05";
            string etdHours = "23";
            string aircraft = TestContext.Properties["Registration"].ToString();
            string validateFormat = "V";
            // arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            //act
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber1, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                flightPage.WaitPageLoading();
                flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                flightPage.WaitPageLoading();
                flightPage.Filter(FilterType.SearchFlight, prefix);
                flightPage.OtherFlights(OtherFlightType.ValidateAll);
                bool check = flightPage.IsValidated();
                //Assert 
                Assert.IsTrue(check, "validateAll ne s'applique pas");
            }
            finally
            {
                //Delete Flight
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(prefix);
                massiveDeleteFlightPage.SelectSiteToFilter(site);
                massiveDeleteFlightPage.SelectCustomer(customer);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.ClickSelectAllButton();
                massiveDeleteFlightPage.Delete();
            }
        }
    }
}