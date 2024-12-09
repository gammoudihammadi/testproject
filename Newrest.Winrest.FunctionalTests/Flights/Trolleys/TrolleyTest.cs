using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Schedule;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Trolleys;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Security.Policy;
using System.Web.Routing;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;

namespace Newrest.Winrest.FunctionalTests.Flights
{
    [TestClass]
    public class TrolleyTest : TestBase
    {
        // Impossible de évaluer l’expression, car un frame natif se trouve en haut de la pile des appels.
        private const int _timeout = 60 * 10 * 1000;

        string flightNumberBonnata = "Bonnata - " + DateUtils.Now.ToString("dd-MM-yyyy");
        string flightNumberConnata = "Connata - " + DateUtils.Now.AddDays(1).ToString("dd-MM-yyyy");
        string flightNumberDarabata = "Darabata - " + DateUtils.Now.AddDays(1).ToString("dd-MM-yyyy");

        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void FL_TR_TrolleyServiceConfig()
        {
            // Prepare Service 1
            string serviceName = TestContext.Properties["TrolleyService"].ToString();
            string serviceCode = TestContext.Properties["CodeTrolley"].ToString();
            string serviceProduction = TestContext.Properties["ServiceProductionTrolley"].ToString();
            string categorie = TestContext.Properties["CategorieBOB"].ToString();
            string guesttype = TestContext.Properties["GuestTypeTrolley"].ToString();
            string customerName = TestContext.Properties["CustomerLP"].ToString();
            string customerNamebis = TestContext.Properties["CustomerNamebis"].ToString();
            string datasheetName = TestContext.Properties["DatasheetName"].ToString();
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            // Prepare Service 2
            string siteMAD = TestContext.Properties["Site"].ToString();
            string serviceName1 = TestContext.Properties["TrolleyService1"].ToString();
            string serviceCode1 = TestContext.Properties["CodeTrolley1"].ToString();
            string serviceProduction1 = TestContext.Properties["ServiceProductionTrolley1"].ToString();
            string guesttype1 = TestContext.Properties["GuestTypeTrolley1"].ToString();
            string customerName1 = TestContext.Properties["customerName1Trolley"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            ClearCache();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            if (servicePage.CheckTotalNumber() == 0)
            {

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction, categorie, guesttype);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerName, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2), "20", datasheetName);
                servicePage = pricePage.BackToList();
            }
            else
            {
                var servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.SearchPriceForCustomer(customerNamebis, siteACE, DateUtils.Now.AddDays(-15), DateUtils.Now.AddMonths(2), "20", datasheetName);

                var serviceGeneralInformationsPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            //Assert
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            if (servicePage.CheckTotalNumber() == 0)
            {

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1, serviceCode1, serviceProduction1, categorie, guesttype1);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteMAD, customerName1, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2), "20", datasheetName1);
                pricePage.BackToList();
            }
            else
            {
                var servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.SearchPriceForCustomer(customerName1, siteMAD, DateUtils.Now.AddDays(-15), DateUtils.Now.AddMonths(2), "20", datasheetName1);

                var serviceGeneralInformationsPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                serviceGeneralInformationsPage.BackToList();
            }

        }

        [TestMethod]
        [Priority(1)]
        [Timeout(_timeout)]
        public void FL_TR_Create_Lp_CardBobConfig()
        {
            // Prepare first LpCart

            string code = TestContext.Properties["CodeTrolley2"].ToString();
            string name = TestContext.Properties["trolleyFlightBob"].ToString();
            string codeBis = TestContext.Properties["CodeTrolley3"].ToString();
            string nameBis = TestContext.Properties["Carav Trole yaeh"].ToString();
            string lpCartCalomato = TestContext.Properties["LpCartName2"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();
            string siteMad = TestContext.Properties["SiteLpCart2"].ToString();
            string customerVLM = TestContext.Properties["customerNameVLMTrolley"].ToString();

            DateTime from = DateUtils.Now.AddDays(-5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Bob comment";


            // Arrange
            var homePage = LogInAsAdmin();


            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartCalomato);

            //On repousse la date du Lpcart
            var LpCartDetailPageForUpdate = lpCartPage.LpCartCartDetailPage();
            var infomationPage = LpCartDetailPageForUpdate.LpCartGeneralInformationPage();
            infomationPage.UpdateToDateTo();
            LpCartDetailPageForUpdate.BackToList();


            // Act
            lpCartPage.Filter(LpCartPage.FilterType.Search, name);

            if (lpCartPage.CheckTotalNumber() == 0)
            {
                // Create first lp cart
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();

                //Acces Onglet Cart1           
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, name, siteMad, customerVLM, "AB310", from, to, comment);

                LpCartDetailPage.ClickAddtrolley();
                LpCartDetailPage.AddTrolley(trolleyName);
                //Create LpCart Scheme
                var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");

                //Acces Onglet Cart2                
                LpCartDetailPage.ClickAddtrolley();
                LpCartDetailPage.AddTrolley(trolleyName);

                //Create LpCart Scheme
                lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");
                LpCartDetailPage.BackToList();
            }
            else
            {
                var LpCartDetailPage = lpCartPage.LpCartCartDetailPage();
                var infoPage = LpCartDetailPage.LpCartGeneralInformationPage();
                infoPage.UpdateToDateTo();
                LpCartDetailPage.BackToList();
            }


            lpCartPage.Filter(LpCartPage.FilterType.Search, nameBis);
            if (lpCartPage.CheckTotalNumber() == 0)
            {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                //Create second LpCart

                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(codeBis, nameBis, siteMad, customerVLM, "B788", from, to, comment);

                //Acces Onglet Cart1                
                LpCartDetailPage.ClickAddtrolley();
                LpCartDetailPage.AddTrolley(trolleyName);
                //Create LpCart Scheme andd position
                var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");

                //Acces Onglet Cart2                
                LpCartDetailPage.ClickAddtrolley();
                LpCartDetailPage.AddTrolley(trolleyName);
                //Create LpCart Scheme
                lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "1", "1");

                //Acces Onglet Cart2                
                LpCartDetailPage.ClickAddtrolley();
                LpCartDetailPage.AddTrolley(trolleyName);
                //Create LpCart Scheme
                lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "1", "1");
                LpCartDetailPage.BackToList();
            }
            else
            {
                var LpCartDetailPage = lpCartPage.LpCartCartDetailPage();
                var infoPage = LpCartDetailPage.LpCartGeneralInformationPage();
                infoPage.UpdateToDateTo();
                LpCartDetailPage.BackToList();
            }


            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            // Assert
            Assert.IsTrue(lpCartPage.CheckTotalNumber() == 1, "Le lpCart n'est pas crée");


            lpCartPage.Filter(LpCartPage.FilterType.Search, nameBis);
            // Assert
            Assert.IsTrue(lpCartPage.CheckTotalNumber() == 1, "Le lpCart n'est pas crée");

        }

        [TestMethod]
        [Priority(2)]
        [Timeout(_timeout)]
        public void FL_TR_Create_LoadingPlan_with_Lp_CardBobConfig()
        {
            // Prepare 2 Lpcart

            string code = TestContext.Properties["CodeTrolley2"].ToString();
            string name = TestContext.Properties["trolleyFlightBob"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();

            string customerLp = TestContext.Properties["customerName1Trolley"].ToString();
            string guestName = TestContext.Properties["GuestNameBob"].ToString();

            string codeBis = TestContext.Properties["CodeTrolley3"].ToString();
            string nameBis = TestContext.Properties["Carav Trole yaeh"].ToString();


            //Prepare Loading plan
            string loadingPlanName = TestContext.Properties["loadingPlanNameTrolley"].ToString();
            string loadingPlanNameBis = TestContext.Properties["loadingPlanNameTrolleyBis"].ToString();
            string aircraft = "AB310";
            string type = TestContext.Properties["TypeBob"].ToString();
            string serviceName1 = TestContext.Properties["TrolleyService1"].ToString();
            string site1 = TestContext.Properties["SiteLpCart2"].ToString();
            string customerName1 = "VLM AIRLINES";


            // Arrange
            var homePage = LogInAsAdmin();


            // Act Loading
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site1);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                // Create first Loading Plan
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, null, aircraft, site1, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //add leg and service bob 
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                if (loadingPlanDetailsPage.IsServiceAdded(serviceName1))
                {
                    loadingPlanDetailsPage.DeleteService();
                    loadingPlanDetailsPage = loadingPlanDetailsPage.GoToLoadingPlanDetailPage();
                }
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName1);

                var generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.SelectLpCart(code + " - " + name);

                loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.ClickLoadingPlanLPCartEditLabels(trolleyScheme);
                generalInformationsPage.BackToList();
            }
            else
            {
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                var loadingPlanDetailsPage = generalInformationsPage.GoToLoadingPlanDetailPage();

                //add leg and service bob 
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                if (loadingPlanDetailsPage.IsServiceAdded(serviceName1))
                {
                    loadingPlanDetailsPage.DeleteService();
                    loadingPlanDetailsPage = loadingPlanDetailsPage.GoToLoadingPlanDetailPage();
                }
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName1);

                generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                if (generalInformationsPage.GetLpCartName() != code + " - " + name)
                {
                    generalInformationsPage.SelectLpCart(code + " - " + name);
                }
                generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.ClickLoadingPlanLPCartEditLabels(trolleyScheme);
                generalInformationsPage.BackToList();
            }

            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site1);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

            // Assert
            Assert.IsTrue(loadingPlanPage.CheckTotalNumber() == 1, "Le loading plan n'est pas crée");

            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site1);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);
            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                ///Second loading plan
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanNameBis, customerName1, null, "B788", site1, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                if (loadingPlanDetailsPage.IsServiceAdded(serviceName1))
                {
                    loadingPlanDetailsPage.DeleteService();
                    loadingPlanDetailsPage = loadingPlanDetailsPage.GoToLoadingPlanDetailPage();
                }
                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName1);
                //Edit the first loading plan
                var generalInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.SelectLpCart(codeBis + " - " + nameBis);
                loadingPlanPage = loadingPlanDetailsPage.BackToList();
            }
            else
            {
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                if (generalInformationsPage.GetLpCartName() != codeBis + " - " + nameBis)
                {
                    generalInformationsPage.SelectLpCart(codeBis + " - " + nameBis);
                }
                var loadingPlanDetailsPage = generalInformationsPage.GoToLoadingPlanDetailPage();
                //add leg and service bob 
                if (loadingPlanDetailsPage.IsServiceAdded(serviceName1))
                {
                    loadingPlanDetailsPage.DeleteService();
                    loadingPlanDetailsPage.WaitPageLoading();
                    loadingPlanDetailsPage = loadingPlanDetailsPage.GoToLoadingPlanDetailPage();
                }
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                if (loadingPlanDetailsPage.IsServiceAdded(serviceName1))
                {
                    loadingPlanDetailsPage.DeleteService();
                    loadingPlanDetailsPage = loadingPlanDetailsPage.GoToLoadingPlanDetailPage();
                }
                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName1);
                generalInformationsPage.BackToList();
            }

            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site1);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);
            // Assert
            Assert.IsTrue(loadingPlanPage.CheckTotalNumber() == 1, "Le loading plan n'est pas crée");
        }



        [TestMethod]
        [Priority(3)]
        [Timeout(_timeout)]
        public void FL_TR_CreateFlight_ForModifTrolleyConfig()
        {
            //prepare

            string siteMad = TestContext.Properties["SiteLpCart2"].ToString();
            string siteAal = TestContext.Properties["SiteAalTrolley"].ToString();
            string customerVLM = TestContext.Properties["customerName1Trolley"].ToString();
            string trolleyFlightBob = TestContext.Properties["trolleyFlightBob"].ToString();
            string trolleyServiceBob = TestContext.Properties["TrolleyService1"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var flightPage = homePage.GoToFlights_FlightPage();

            // Create
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteMad);
            flightPage.SetDateState(DateUtils.Now.AddDays(1));
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberDarabata);

            if (flightPage.CheckTotalNumber() == 0)
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumberDarabata.ToString(), customerVLM, "AB310", siteMad, siteAal, null, "13", "15", trolleyFlightBob, DateUtils.Now.AddDays(1));
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberDarabata);
                var editPage = flightPage.EditFirstFlight(flightNumberDarabata);
                editPage.AddGuestType("BOB");
                editPage.AddService(trolleyServiceBob);
                editPage.CloseViewDetails();
            }
            else
            {
                var editPage = flightPage.EditFirstFlight(flightNumberDarabata);
                editPage.SelectLpCart(trolleyFlightBob);
                editPage.CloseViewDetails();
            }

            //Assert
            Assert.AreEqual(flightNumberDarabata.ToString(), flightPage.GetFirstFlightNumber(), "Le vol n'a pas été créé.");
        }

        [TestMethod]
        [Priority(5)]
        [Timeout(_timeout)]
        public void FL_TR_Flight_TrolleyAdjustingConfig()
        {
            string lpcart = TestContext.Properties["LpCartName2"].ToString(); // Calomato : Site MAD Customer "AEE - AEGEAN AIRLINES" Aircraft "0FICIN"
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString(); // 0FICIN
            string customer = TestContext.Properties["CustomerLpCart2"].ToString(); // AEE - AEGEAN AIRLINES
            string ACE = TestContext.Properties["Site"].ToString(); // MAD
            string service = "HOT BREAKFAST BC AEE";

            //Arrange
            var homePage = LogInAsAdmin();


            //Prepare
            // Prepare
            string deliveryCustomer = TestContext.Properties["CustomerLpCart2"].ToString();
            string deliverySite = TestContext.Properties["Site"].ToString();

            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);

            // vérification du service
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(service);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(deliverySite, deliveryCustomer, fromDate, toDate);
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.UnfoldAll();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(deliverySite, deliveryCustomer);
                serviceCreatePriceModalPage.EditPriceDates(fromDate, toDate);
                if (!serviceCreatePriceModalPage.VerifySuccessEditPriceDates())
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
            }

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberBonnata);
            if (flightPage.IsFlightExist())
            {

                if (flightPage.IsValidated())
                {
                    flightPage.UnSetNewState("V");
                }

                var flightModalEdit = flightPage.EditFirstFlight(flightNumberBonnata);
                flightModalEdit.SelectLpCart(lpcart);
                flightModalEdit.CloseViewDetails();
            }
            else
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumberBonnata.ToString(), customer, aircraft, "MAD", ACE, null, "13", "18", lpcart);
            }
            var editPage = flightPage.EditFirstFlight(flightNumberBonnata);

            if (editPage.CountServicesFromBob() == 1 || editPage.IsGuestTypeVisible())
            {
                editPage.DeleteGuestType();
            }

            editPage.AddGuestType("BOB");
            editPage.AddService(service);

            homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberConnata);
            flightPage.SetDateState(DateUtils.Now.AddDays(1));

            if (flightPage.IsFlightExist())
            {
                if (flightPage.IsValidated())
                {
                    flightPage.UnSetNewState("V");
                }
                var flightModalEdit = flightPage.EditFirstFlight(flightNumberConnata);
                flightModalEdit.SelectLpCart(lpcart);
                flightModalEdit.CloseViewDetails();
            }
            else
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumberConnata.ToString(), customer, aircraft, "MAD", ACE, null, "13", "18", lpcart, DateUtils.Now.AddDays(1));
            }
            flightPage.EditFirstFlight(flightNumberConnata);

            if (editPage.CountServicesFromBob() == 1 || editPage.IsGuestTypeVisible())
            {
                editPage.DeleteGuestType();
            }

            editPage.AddGuestType("BOB");
            editPage.AddService(service);

            // ajout de la position et de la quantité
            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();

            trolleyLightLabelPage.ResetFilters();
            var trolleyDetailledPage = trolleyLightLabelPage.GoToDetailledLabel();
            trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));
            trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.Site, "MAD");
            trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.SearchFlight, flightNumberBonnata);
            trolleyDetailledPage.AddPositionAndQuantity("accotroScheme", "2");
            trolleyLightLabelPage.ResetFilters();
            trolleyDetailledPage = trolleyLightLabelPage.GoToDetailledLabel();
            trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.ProdDate, DateUtils.Now);
            trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.Site, "MAD");
            trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.SearchFlight, flightNumberConnata);
            trolleyDetailledPage.AddPositionAndQuantity("accotroScheme", "2");

            //retour sur Flights et clique sur parcette du vol
            homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberConnata);
            flightPage.SetDateState(DateUtils.Now.AddDays(1));
            flightPage.UnSetNewState("P");
            flightPage.SetNewState("P");
            flightPage.ClickFirstFlightKey();

            homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberBonnata);
            flightPage.UnSetNewState("P");
            flightPage.SetNewState("P");
            flightPage.ClickFirstFlightKey();
        }

        [TestMethod]
        [Priority(6)]
        [Timeout(_timeout)]
        public void FL_TR_AdjustQuantity_TrolleyAdjusting()
        {

            string position = "accotroScheme";
            string qty = 2.ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            // Ancien  Flight flightNumberConnata, parfois flightNumberDarabata
            string flightNumber;

            flightNumber = flightNumberDarabata;

            ////Act
            ///
            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();
            var trolleyDetailedLabelPage = trolleyLightLabelPage.GoToDetailledLabel();

            trolleyDetailedLabelPage.Filter(TrolleyDetailedLabelPage.FilterType.ProdDate, DateUtils.Now);
            trolleyDetailedLabelPage.Filter(TrolleyDetailedLabelPage.FilterType.Site, "MAD");
            trolleyDetailedLabelPage.Filter(TrolleyDetailedLabelPage.FilterType.SearchFlight, flightNumber);
            if (trolleyDetailedLabelPage.CheckTotalNumber() == 0)
            {
                flightNumber = flightNumberConnata;
                trolleyDetailedLabelPage.Filter(TrolleyDetailedLabelPage.FilterType.SearchFlight, flightNumber);
            }

            trolleyDetailedLabelPage.AddPositionAndQuantity(position, qty);

            Assert.IsTrue(trolleyDetailedLabelPage.verifyPosition(position), "l'ajout de position n'est pas fonctionner");
            Assert.IsTrue(trolleyDetailedLabelPage.verifyQty(qty), "l'ajout de position n'est pas fonctionner");


            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.SetDateState(DateUtils.Now.AddDays(1));

            //décocher et Cocher pour appliquer ajustement          
            if (flightPage.IsPreval())
            {
                flightPage.UnSetNewState("P");
            }
            flightPage.SetNewState("P");

            trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();
            var trolleyAjustingPage = trolleyLightLabelPage.GoToTrolleyAjusting();

            trolleyAjustingPage.Filter(TrolleyAdjustingPage.FilterType.ProdDate, DateUtils.Now);
            trolleyAjustingPage.Filter(TrolleyAdjustingPage.FilterType.Site, "MAD");
            trolleyDetailedLabelPage.Filter(TrolleyDetailedLabelPage.FilterType.SearchFlight, flightNumber);

            //click ajust 
            trolleyAjustingPage.ClickAjust();

            //Assert
            Assert.IsTrue(trolleyAjustingPage.IsAjusted(), "Le trolley n'est pas ajusté.");

        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_Filter_Service()
        {
            //Prepare
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "ACE");

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType("BOB");
            editPage.AddService("Trolley Service");
            editPage.CloseViewDetails();


            //Second Flight
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");

            flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, "VLM AIRLINES", aircraft, "MAD", siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType("BOB");
            editPage.AddService("Bob Service for trolley");
            editPage.CloseViewDetails();

            // Flight>Schedules
            // D-1 => D
            SchedulePage schedule = homePage.GoToFlights_FlightSelectionPage();
            if (!schedule.isPageSizeEqualsTo100())
            {
                schedule.PageSize("8");
                schedule.PageSize("100");
            }
            schedule.Filter(SchedulePage.FilterType.Site, "MAD");
            schedule.UnfoldAll();
            schedule.SetFlightProduced(flightNumber, true);

            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();
            trolleyLightLabelPage.ResetFilters();

            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.SearchService, "Trolley Service");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, "ACE");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));

            // Assert
            Assert.IsTrue(trolleyLightLabelPage.VerifyService("Trolley Service"), MessageErreur.FILTRE_ERRONE, "Services");

            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.SearchService, "Bob Service for trolley");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, "MAD");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));

            // Assert
            Assert.IsTrue(trolleyLightLabelPage.VerifyService("Bob Service for trolley"), MessageErreur.FILTRE_ERRONE, "Services");

        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_Filter_Flight()
        {
            //Prepare
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "ACE");

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType("BOB");
            editPage.AddService("Trolley Service");
            editPage.CloseViewDetails();

            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();
            trolleyLightLabelPage.ResetFilters();

            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.SearchFlight, flightNumber);
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, "ACE");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));

            // Assert
            Assert.IsTrue(trolleyLightLabelPage.VerifyFlight(flightNumber), MessageErreur.FILTRE_ERRONE, "Flight");

        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_Filter_Site()
        {
            //Prepare
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "ACE");

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType("BOB");
            editPage.AddService("Trolley Service");
            editPage.CloseViewDetails();

            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();
            trolleyLightLabelPage.ResetFilters();



            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.SearchService, "Trolley Service");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, "ACE");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));

            if (!trolleyLightLabelPage.isPageSizeEqualsTo100())
            {
                trolleyLightLabelPage.PageSize("8");
                trolleyLightLabelPage.PageSize("100");
            }

            // Assert
            Assert.IsTrue(trolleyLightLabelPage.VerifySite("ACE"), MessageErreur.FILTRE_ERRONE, "Services");

        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_Filter_ProdDate()
        {
            //Prepare
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["Site"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();
            string dateFormat = homePage.GetDateFormatPickerValue();



            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType("BOB");
            editPage.AddService("Trolley Service");
            editPage.CloseViewDetails();

            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();

            trolleyLightLabelPage.ResetFilters();
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, siteFrom);
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));

            if (!trolleyLightLabelPage.isPageSizeEqualsTo100())
            {
                trolleyLightLabelPage.PageSize("8");
                trolleyLightLabelPage.PageSize("100");
            }

            // Assert
            Assert.IsTrue(trolleyLightLabelPage.IsDateRespected(DateUtils.Now.AddDays(-1), dateFormat), MessageErreur.FILTRE_ERRONE, "Services");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TR_Filter_SortByFlightNumber()
        {
            // prepare 
            string serviceName1 = "service_1_" + DateUtils.Now;
            string site = "ACE";
            string loadingPlanName = new Random().Next().ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string flightNumber1 = "flight_lp1" + new Random().Next(1, 5000).ToString();
            string flightNumber2 = "flight_lp2" + new Random().Next(5001, 10000).ToString();
            DateTime serviceDateFrom = DateUtils.Now;
            DateTime serviceDateTo = DateUtils.Now.AddMonths(2);
            DateTime lpDateFrom = serviceDateFrom.AddDays(2);
            DateTime lpDateTo = serviceDateTo.AddDays(-2);
            string  lpCartName = "lpcart bob" + new Random().Next().ToString();
            Random rnd = new Random();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string type = "BuyOnBoard";
            string guestName = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();
            string code = "Bob" + DateUtils.Now.ToString("dd/MM/yyyy") + rnd.Next().ToString();
            string comment = "Bob comment";
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string etaHours = "00";
            string etdHours = "23";
            string flightNumber = "FlightNo";
            // arrange
            var homePage = LogInAsAdmin();
            // act
            try
            {
                var servicePage = homePage.GoToCustomers_ServicePage();
                var ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName1);
                var serviceGeneralInformationsPage = ServiceCreatePage.Create();
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, serviceDateFrom, serviceDateTo);
                servicePage = pricePage.BackToList();
                // crée lpcart
                var lpCartPage = homePage.GoToFlights_LpCartPage();
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCartWithRoutes(code, lpCartName, site, customer, aircraft, lpDateFrom, lpDateTo, comment, route);
                LpCartDetailPage.BackToList();
                // Create LoadingPlan
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customer, route, aircraft, site, type);
                loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(lpDateTo, lpDateFrom);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                loadingPlanDetailsPage.ClickGuestBtn();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName1);
                loadingPlanDetailsPage.ClickOnDetailsPage();
                loadingPlanDetailsPage.BackToList();
                //Create Flight
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber1, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, lpDateFrom);
                flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, lpDateFrom);
                var trolleyPage = flightPage.GoToFlights_TrolleyPage();
                trolleyPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, lpDateFrom.AddDays(-1));
                trolleyPage.WaitPageLoading();
                trolleyPage.Filter(TrolleyLightLabelPage.FilterType.SortBy, flightNumber);
                bool check = trolleyPage.IsSortedByFlight();
                Assert.IsTrue(check,"le filtre sortby ne s'applique pas");
            }
            finally
            {
                var flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber1);
                massiveDeleteFlightPage.SetDateFromAndTo(lpDateFrom, lpDateTo);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
                flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber2);
                massiveDeleteFlightPage.SetDateFromAndTo(lpDateFrom, lpDateTo);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
                //Delete LoadingPlan
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, null, null, lpDateTo.AddMonths(2));
                //Delete LpCart
                var lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                lpCartPage.DeleteLpCart();
                var servicePage = homePage.GoToCustomers_ServicePage();
                //Delete service1
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName1);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TR_Filter_SortByServiceName()
        {
            // prepare 
            string serviceName2 = "service_2_" + DateUtils.Now;
            string serviceName1 = "service_1_" + DateUtils.Now;
            string site = "ACE";
            string loadingPlanName = new Random().Next().ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string flightNumber1 = "flight_lp1" + new Random().Next(1, 5000).ToString();
            string flightNumber2 = "flight_lp2" + new Random().Next(5001, 10000).ToString();
            DateTime serviceDateFrom = DateUtils.Now;
            DateTime serviceDateTo = DateUtils.Now.AddMonths(2);
            DateTime lpDateFrom = serviceDateFrom.AddDays(2);
            DateTime lpDateTo = serviceDateTo.AddDays(-2);
            string lpCartName = "lpcart bob" + new Random().Next().ToString();
            Random rnd = new Random();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string type = "BuyOnBoard";
            string guestName = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();
            string code = "Bob" + DateUtils.Now.ToString("dd/MM/yyyy") + rnd.Next().ToString();
            string comment = "Bob comment";
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string etaHours = "00";
            string etdHours = "23";
            string ServiceName = "ServiceName";
            // arrange
            var homePage = LogInAsAdmin();
            // act
            try
            {
                // crée service 1
                var servicePage = homePage.GoToCustomers_ServicePage();
                var ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName1);
                var serviceGeneralInformationsPage = ServiceCreatePage.Create();
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, serviceDateFrom, serviceDateTo);
                servicePage = pricePage.BackToList();
                // crée service 2 
                ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName2);
                serviceGeneralInformationsPage = ServiceCreatePage.Create();
                pricePage = serviceGeneralInformationsPage.GoToPricePage();
                priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, serviceDateFrom, serviceDateTo);
                servicePage = pricePage.BackToList();
                // crée lpcart
                var lpCartPage = homePage.GoToFlights_LpCartPage();
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCartWithRoutes(code, lpCartName, site, customer, aircraft, lpDateFrom, lpDateTo, comment, route);
                LpCartDetailPage.BackToList();
                // Create LoadingPlan
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customer, route, aircraft, site, type);
                loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(lpDateTo, lpDateFrom);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                loadingPlanDetailsPage.ClickGuestBtn();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName1);
                loadingPlanDetailsPage.ClickOnDetailsPage();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName2);
                loadingPlanDetailsPage.BackToList();
                //Create Flight
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber1, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, lpDateFrom);
                flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, lpDateFrom);
                var trolleyPage = flightPage.GoToFlights_TrolleyPage();
                trolleyPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, lpDateFrom.AddDays(-1));
                trolleyPage.WaitPageLoading();
                trolleyPage.Filter(TrolleyLightLabelPage.FilterType.SortBy, ServiceName);
                bool check = trolleyPage.IsSortedByService();
                Assert.IsTrue(check, "le filtre sortby ne s'applique pas");
            }
            finally
            {
                var flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber1);
                massiveDeleteFlightPage.SetDateFromAndTo(lpDateFrom, lpDateTo);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
                flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber2);
                massiveDeleteFlightPage.SetDateFromAndTo(lpDateFrom, lpDateTo);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
                //Delete LoadingPlan
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, null, null, lpDateTo.AddMonths(2));
                //Delete LpCart
                var lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                lpCartPage.DeleteLpCart();
                var servicePage = homePage.GoToCustomers_ServicePage();
                //Delete service1
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName1);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName2);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_Filter_Customers()
        {
            //Prepare
            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, "VLM AIRLINES", aircraft, "MAD", siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType("BOB");
            editPage.AddService("Bob Service for trolley");
            editPage.CloseViewDetails();

            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();
            trolleyLightLabelPage.ResetFilters();

            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, "MAD");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Customers, "VLM - VLM AIRLINES");

            if (!trolleyLightLabelPage.isPageSizeEqualsTo100())
            {
                trolleyLightLabelPage.PageSize("8");
                trolleyLightLabelPage.PageSize("100");
            }

            // Assert
            Assert.IsTrue(trolleyLightLabelPage.IsCustomer("VLM AIRLINES"), MessageErreur.FILTRE_ERRONE, "Customer");


        }


        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_Filter_GuestTypes()
        {
            //Prepare
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "ACE");

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType("BOB");
            editPage.AddService("Trolley Service");
            editPage.CloseViewDetails();

            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();
            trolleyLightLabelPage.ResetFilters();

            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, "ACE");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.GuestTypes, "BOB");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.SearchFlight, flightNumber);

            var firstFlight = trolleyLightLabelPage.GetFirstFlight();

            homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, firstFlight);
            flightPage.Filter(FlightPage.FilterType.Sites, "ACE");

            var detailFlight = flightPage.EditFirstFlight(firstFlight);
            var guestType = detailFlight.GetGuestType();

            Assert.AreEqual("BOB", guestType, MessageErreur.FILTRE_ERRONE, "Guest Type");
        }


        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_PrintLabel()
        {
            //FIXME
            //fonctionne si on lance FL_TR_PrintDetailledLabel avant

            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();


            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();

            trolleyLightLabelPage.ClearDownloads();

            trolleyLightLabelPage.ResetFilters();
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));

            // Impression
            var reportPage = trolleyLightLabelPage.PrintReport(TrolleyLightLabelPage.PrintType.PrintLabel, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_PrintLabelWithHideLabel()
        {
            //Prepare
            string flightNumber = new Random().Next().ToString();
            string customer = "SMARTWINGS, A.S. (TVS)";
            string aircraft = "B737";
            string guest = "YC";
            string site = "ACE";
            //Arrange
            var homePage = LogInAsAdmin();



            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, site);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, site);

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType(guest);
            editPage.AddService("Trolley Service");
            editPage.CloseViewDetails();

            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();

            trolleyLightLabelPage.ClearDownloads();


            trolleyLightLabelPage.ResetFilters();
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, site);
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.SearchService, "Trolley Service");

            var firstFlightItem = trolleyLightLabelPage.GetFirstFlight();
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.SearchFlight, firstFlightItem);
            trolleyLightLabelPage.ClickHideLabel();

            // Impression
            var reportPage = trolleyLightLabelPage.PrintReportErrorConfig(TrolleyLightLabelPage.PrintType.PrintLabel);

            Assert.IsTrue(reportPage, "Le message d'erreur ne s'affiche pas.");
        }


        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_PrintLabelWithHighQty()
        {
            //Arrange
            var homePage = LogInAsAdmin();


            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();

            trolleyLightLabelPage.ClearDownloads();

            trolleyLightLabelPage.ResetFilters();
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, "ACE");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.SearchService, "Trolley Service");

            var firstFlightItem = trolleyLightLabelPage.GetFirstFlight();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, firstFlightItem);
            flightPage.Filter(FlightPage.FilterType.Sites, "ACE");

            var detailFlight = flightPage.EditFirstFlight(firstFlightItem);
            detailFlight.SetFinalQty("1000");

            homePage.GoToFlights_TrolleyPage();
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));

            // Impression
            var reportPage = trolleyLightLabelPage.PrintReportErrorConfig(TrolleyLightLabelPage.PrintType.PrintLabel);

            homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, firstFlightItem);
            flightPage.Filter(FlightPage.FilterType.Sites, "ACE");
            flightPage.EditFirstFlight(firstFlightItem);
            detailFlight.SetFinalQty("1");

            Assert.IsTrue(reportPage, "Le message d'erreur ne s'affiche pas.");
        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_PrintLabelNotConfig()
        {
            //Arrange
            var homePage = LogInAsAdmin();


            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();

            trolleyLightLabelPage.ClearDownloads();

            trolleyLightLabelPage.ResetFilters();
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, "ACE");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.SearchService, "Service for BOB");

            // Impression
            var reportPage = trolleyLightLabelPage.PrintReportErrorConfig(TrolleyLightLabelPage.PrintType.PrintLabel);

            Assert.IsTrue(reportPage, "Le message d'erreur ne s'affiche pas.");
        }


        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_SetConfigOnLightLabel()
        {
            string trolleyServiceBob = TestContext.Properties["TrolleyService1"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();
            trolleyLightLabelPage.ResetFilters();
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, "MAD");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.SearchService, trolleyServiceBob);

            // Impression
            var configModalPage = trolleyLightLabelPage.ConfigurationModalPage();

            configModalPage.SetValuesForService("1");

            Assert.IsTrue(trolleyLightLabelPage.IsConfigIsSet(), "La configuration ne sait pas appliqué.");

            trolleyLightLabelPage.ConfigurationModalPage();
            configModalPage.DeleteValuesForService();

            Assert.IsTrue(trolleyLightLabelPage.IsConfigIsDeleted(), "La configuration ne sait pas appliqué.");
        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_PrintDetailledLabel()
        {
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLpCart2"].ToString();
            string ACE = TestContext.Properties["Site"].ToString();
            string lpcart = TestContext.Properties["LpCartName2"].ToString();
            var flightNumber = "Connata - " + DateUtils.Now.ToString("dd-MM-yyyy");

            bool newVersionPrint = true;
            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);

            if (!flightPage.IsFlightExist())
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber.ToString(), customer, aircraft, "MAD", ACE, null, "13", "18", lpcart);
            }

            if (flightPage.IsValidated())
            {
                flightPage.UnSetNewState("V");
            }

            var editPage = flightPage.EditFirstFlight(flightNumber);

            if (editPage.CountServicesFromBob() == 1 || !editPage.GetGuestType().Equals(""))
            {
                editPage.DeleteGuestType();
            }

            editPage.AddGuestType("BOB");
            editPage.AddService("HOT BREAKFAST BC AEE");

            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();

            trolleyLightLabelPage.ClearDownloads();

            trolleyLightLabelPage.ResetFilters();
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.Site, "MAD");
            trolleyLightLabelPage.Filter(TrolleyLightLabelPage.FilterType.ProdDate, DateTime.Now.AddDays(-1));

            var trolleyDetailledPage = trolleyLightLabelPage.GoToDetailledLabel();

            Assert.IsTrue(trolleyDetailledPage.CheckTotalNumber() > 0, "Aucun item est dans le Detailed label.");

            // Impression
            var reportPage = trolleyDetailledPage.PrintReport(TrolleyDetailedLabelPage.PrintType.Print, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
            //bug il manque la reherrche de service sur le pdf 

        }


        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_Filter_Search_FlightKey_TrolleyAdjusting()
        {


            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberBonnata);

            var flightKey = flightPage.GetFirstFlightKey();

            var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();
            var trolleyAjustingPage = trolleyLightLabelPage.GoToTrolleyAjusting();

            trolleyAjustingPage.Filter(TrolleyAdjustingPage.FilterType.ProdDate, DateUtils.Now.AddDays(-1));
            trolleyAjustingPage.Filter(TrolleyAdjustingPage.FilterType.Site, "MAD");
            trolleyAjustingPage.Filter(TrolleyAdjustingPage.FilterType.SearchFlightKey, flightKey);

            //Assert
            Assert.IsTrue(trolleyAjustingPage.VerifyFlight(flightNumberBonnata), "Le filtre Flight Key dans les trolley ajusting ne marche pas.");

        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_CountedQuantity_TrolleyCount()
        {

            string flightNumber = flightNumberDarabata;
            //Arrange
            var homePage = LogInAsAdmin();


            ////Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.Filter(FlightPage.FilterType.Status, "Preval");
            flightPage.SetDateState(DateUtils.Now.AddDays(1));
            if (flightPage.CheckTotalNumber() == 0)
            {
                flightNumber = flightNumberConnata;
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            }



            //Cocher et décocher pour appliquer ajustement  
            flightPage.UnSetNewState("P");
            flightPage.SetNewState("P");
            flightPage.SetNewState("V");
            var trolleyAjustingPage = flightPage.ClickFirstFlightKey();

            var trolleyCountingPage = trolleyAjustingPage.GoToTrolleyCounting();

            //click ajust 
            trolleyCountingPage.ClickCount();

            //Assert
            Assert.IsTrue(trolleyCountingPage.IsCounted(), "L'application n'a pas pu générer le fichier attendu.");
        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_AdjustQuantity_TrolleyCount()
        {


            string trolleyFlightBob = TestContext.Properties["trolleyFlightBob"].ToString();
            string flightNumber = flightNumberDarabata;
            //Arrange
            var homePage = LogInAsAdmin();


            ////Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            // flightNumberDarabata déjà utilisé par FL_TR_ModifyTrolleyOnFlightAjusted 
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.Filter(FlightPage.FilterType.Status, "Preval");
            flightPage.SetDateState(DateUtils.Now.AddDays(1));
            if (flightPage.CheckTotalNumber() == 0)
            {
                flightNumber = flightNumberConnata;
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            }

            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.SelectLpCart(trolleyFlightBob);
            editPage.CloseViewDetails();
            if (!flightPage.IsValidated())
            {
                flightPage.SetNewState("V");
            }

            var trolleyAjustingPage = flightPage.ClickFirstFlightKey();

            var trolleyCountingPage = trolleyAjustingPage.GoToTrolleyCounting();

            //click ajust 
            var qty = trolleyCountingPage.ClickCount();

            //Assert
            Assert.IsTrue(qty, "La quantité ajusté n'est pas ajusté");
            Assert.IsTrue(trolleyCountingPage.IsCounted(), "Le trolley count n'est pas coché");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TR_ModifyTrolleyOnFlightAjusted()
        {
            //bien vérifier que service et customer sont en BOB, flight    
            bool newVersionPrint = true;
            string siteMad = TestContext.Properties["SiteLpCart2"].ToString();
            string siteAal = TestContext.Properties["SiteAalTrolley"].ToString();
            string customerVLM = TestContext.Properties["customerName1Trolley"].ToString();
            string trolleyFlightBob = TestContext.Properties["trolleyFlightBob"].ToString();
            string trolleyServiceBob = TestContext.Properties["TrolleyService1"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            var flightPage = homePage.GoToFlights_FlightPage();

            // Create
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteMad);
            flightPage.SetDateState(DateUtils.Now.AddDays(1));
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberDarabata);

            if (flightPage.CheckTotalNumber() == 0)
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumberDarabata.ToString(), customerVLM, "AB310", siteMad, siteAal, null, "13", "15", trolleyFlightBob, DateUtils.Now.AddDays(1));
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberDarabata);
                var editPage = flightPage.EditFirstFlight(flightNumberDarabata);
                editPage.AddGuestType("BOB");
                editPage.AddService(trolleyServiceBob);
                editPage.SelectLpCart(trolleyFlightBob);
                editPage.CloseViewDetails();
            }
            else
            {
                var editPage = flightPage.EditFirstFlight(flightNumberDarabata);
                if (editPage.GetFlightStatus("V"))
                {
                    editPage.UnsetState("V");
                }
                editPage.ChangeAircaftForLoadingPlan("AB310", false);
                editPage.CloseDataAlertCancel();
                var flightEditModalPage = editPage.EditFlightInModal();
                flightEditModalPage.SelectAllLoadingPlan();
                flightEditModalPage.IsLpCartsDetectedInSelectedLoadingPlan();
                editPage.CloseViewDetails();
            }

            //Assert
            Assert.AreEqual(flightNumberDarabata.ToString(), flightPage.GetFirstFlightNumber(), "Le vol n'a pas été créé.");

            try
            {
                //Act
                var trolleyLightLabelPage = homePage.GoToFlights_TrolleyPage();

                if (newVersionPrint)
                {
                    trolleyLightLabelPage.ClearDownloads();
                }

                var trolleyDetailledPage = trolleyLightLabelPage.GoToDetailledLabel();

                trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.ProdDate, DateUtils.Now);
                trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.Site, siteMad);
                trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.SearchFlight, flightNumberDarabata);

                trolleyDetailledPage.AddPositionAndQuantity("accotroScheme", "2");

                //Act
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteMad);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberDarabata);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));

                var editFlight = flightPage.EditFirstFlight(flightNumberDarabata);
                editFlight.SetFinalQty("10");
                //Cocher et décocher pour appliqer ajustement
                if (editFlight.GetFlightStatus("P"))
                {
                    editFlight.UnsetState("P");
                }
                editFlight.SetNewState("P");
                editFlight.CloseModal();
                var trolleyAjustedPage = flightPage.ClickFirstFlightKey();

                //Assert
                Assert.IsTrue(trolleyAjustedPage.CheckTotalNumber() == 2, "L'application n'affiche pas les trolleys selon l'ajustement des qty dans le detail label");
                Assert.IsTrue(trolleyAjustedPage.IsAllQuantity(), "L'application n'affiche pas les trolleys selon l'ajustement des qty dans le detail label.");

                ////Act
                homePage.GoToFlights_FlightPage();

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteMad);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberDarabata);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                var editFlightPage = flightPage.EditFirstFlight(flightNumberDarabata);
                editFlightPage.SelectLpCart("None");


                editFlightPage.ChangeAircaftForLoadingPlan("B788", false);
                editFlightPage.CloseDataAlertCancel();
                var flightEditModalPage = editFlightPage.EditFlightInModal();
                flightEditModalPage.SelectAllLoadingPlan();
                //editFlightPage.OkBtn();
                editFlightPage.CloseViewDetails();
                editFlightPage.Close();

                //Cocher et décocher pour appliquer ajustement  
                flightPage.UnSetNewState("P");
                flightPage.SetNewState("P");

                //Act
                homePage.GoToFlights_TrolleyPage();

                if (newVersionPrint)
                {
                    trolleyLightLabelPage.ClearDownloads();
                }

                trolleyLightLabelPage.GoToDetailledLabel();

                trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.ProdDate, DateUtils.Now);
                trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.Site, siteMad);
                trolleyDetailledPage.Filter(TrolleyDetailedLabelPage.FilterType.SearchFlight, flightNumberDarabata);
                trolleyDetailledPage.AddPositionAndQuantity("accotroScheme", "2");

                //Act
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteMad);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberDarabata);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));

                editFlight = flightPage.EditFirstFlight(flightNumberDarabata);
                editFlight.SetFinalQty("10");
                //Cocher et décocher pour appliquer ajustement
                if (editFlight.GetFlightStatus("P"))
                {
                    editFlight.UnsetState("P");
                }
                editFlight.SetNewState("P");
                editFlight.CloseModal();

                flightPage.ClickFirstFlightKey();

                //Assert
                //Assert.AreEqual(2, trolleyAjustedPage.CheckTotalNumber(), "L'application n'affiche pas les trolleys selon l'ajustement des qty dans le detail label");
                Assert.IsTrue(trolleyAjustedPage.IsAllQuantity(), "L'application n'affiche pas les trolleys selon l'ajustement des qty dans le detail label.");
            }
            finally
            {
                //Act
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.Filter(FlightPage.FilterType.Sites, siteMad);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberDarabata);

                var editFlightPage = flightPage.EditFirstFlight(flightNumberDarabata);
                editFlightPage.SelectLpCart("None");

                editFlightPage.ChangeAircaftForLoadingPlan("AB310", false);
            }
        }

        [TestMethod]

        [Timeout(_timeout)]
        public void FL_TR_Trolley_Detailed_Label_Print_simplified()
        {
            //Prepare
            string Customer = TestContext.Properties["CustomerLP"].ToString();
            string CustomerFilter = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string categorie = TestContext.Properties["CategorieBOB"].ToString();
            string guestType = TestContext.Properties["GuestNameBob"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            string serviceName = "ServiceForBOBLPCart";
            string LPCartName = "LPCartBOBPrint";
            string BarsetName = "BarsetsTest" + new Random().Next().ToString();
            DateTime dateFrom = DateUtils.Now;
            DateTime dateTo = DateUtils.Now.AddMonths(1);
            string code = new Random().Next().ToString();
            bool newVersionPrint = true;
            LpCartCartDetailPage lpCartCartDetailPage = new LpCartCartDetailPage(WebDriver, TestContext);

            //Arrange            
            var homePage = LogInAsAdmin();

            //Act
            var CustomerPage = homePage.GoToCustomers_CustomerPage();
            CustomerPage.Filter(PageObjects.Customers.Customer.CustomerPage.FilterType.Search, Customer);
            CustomerGeneralInformationPage customerGeneralInformationPage = CustomerPage.SelectFirstCustomer();
            var isBOBChecked = customerGeneralInformationPage.isBuyOnBoard();
            if (!isBOBChecked)
            {
                customerGeneralInformationPage.ActiveBuyOnBoard();
            }
            ServicePage servicePage = customerGeneralInformationPage.GoToCustomers_ServicePage();
            servicePage.Filter(ServicePage.FilterType.Customers, CustomerFilter);
            servicePage.Filter(ServicePage.FilterType.Sites, site);
            servicePage.Filter(ServicePage.FilterType.Categories, categorie);
            if (servicePage.CheckTotalNumber() == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, code, null, categorie, guestType);
                ServiceGeneralInformationPage serviceGeneralInformationPage = serviceCreateModalPage.Create();
                ServicePricePage servicePricePage = serviceGeneralInformationPage.GoToPricePage();
                ServiceCreatePriceModalPage serviceCreatePriceModalPage = servicePricePage.AddNewCustomerPrice();
                serviceCreatePriceModalPage.FillFields_CustomerPrice(site, Customer, dateFrom, dateTo, "20");
            }
            var LPCartPage = homePage.GoToFlights_LpCartPage();
            LPCartPage.Filter(LpCartPage.FilterType.Search, LPCartName);

            if (LPCartPage.CheckTotalNumber() == 0)
            {
                LpCartCreateModalPage lpCartCreateModalPage = LPCartPage.LpCartCreatePage();
                lpCartCartDetailPage = lpCartCreateModalPage.FillField_CreateNewLpCart(LPCartName, LPCartName, site, CustomerFilter, aircraft, dateFrom, dateTo, "");
            }
            else
            {
                lpCartCartDetailPage = LPCartPage.ClickFirstLpCart();
                LpCartGeneralInformationPage lpCartGeneralInformationPage = lpCartCartDetailPage.LpCartGeneralInformationPage();
                lpCartGeneralInformationPage.UpdateFromToDate(dateFrom, dateTo);
                lpCartCartDetailPage = lpCartGeneralInformationPage.ClickOnCarts();
            }

            if (lpCartCartDetailPage.CheckTotalNumber() == 0)
            {
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.CreateNewRowLpCart(1, "BOBTrolley", "BOBTrolleyCode", "BOBTrolleyLocation", "", "", "", "");
                WebDriver.Navigate().Refresh();
                LpCartEditLpCartSchemeModal lpCartEditLpCartSchemeModal = lpCartCartDetailPage.EditCartScheme();
                lpCartEditLpCartSchemeModal.EditLpCartscheme("BobSchema", "2", "1", "1", "1", "blue");
            }
            else
            {
                lpCartCartDetailPage.DeleteAllLpCartScheme();
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.CreateNewRowLpCart(1, "BOBTrolley", "BOBTrolleyCode", "", "", "", "", "");
                WebDriver.Navigate().Refresh();
                LpCartEditLpCartSchemeModal lpCartEditLpCartSchemeModal = lpCartCartDetailPage.EditCartScheme();
                lpCartEditLpCartSchemeModal.EditLpCartscheme("BobSchema", "2", "1", "1", "1", "blue");

            }

            var flightPage = homePage.GoToFlights_FlightPage();
            BarsetsPage barsetsPage = flightPage.ClickOnTop();
            BarsetMultiCreateModalPage barsetMultiCreateModalPage = barsetsPage.CreateMultiBarsets();
            barsetMultiCreateModalPage.FillField_CreateMultiBarsets(site, BarsetName, aircraft, dateFrom, dateTo, CustomerFilter, LPCartName);
            barsetMultiCreateModalPage.CreateButton();
            barsetsPage = flightPage.ClickOnTop();
            barsetsPage.Filter(BarsetsPage.FilterType.SearchFlight, BarsetName);
            barsetsPage.Filter(BarsetsPage.FilterType.Customer, CustomerFilter);
            barsetsPage.Filter(BarsetsPage.FilterType.Sites, site);
            BarsetEditModalPage barsetEditModalPage = barsetsPage.EditFirstBarset();
            barsetEditModalPage.AddGuestType(guestType);
            barsetEditModalPage.AddFirstServiceBarset();
            barsetsPage = barsetEditModalPage.CloseViewDetails();
            TrolleyAdjustingPage trolleyAdjustingPage = barsetsPage.ClickOnNumBarset();
            TrolleyDetailedLabelPage trolleyDetailedLabelPage = trolleyAdjustingPage.GoToTrolleyDetailedLabel();
            if (trolleyDetailedLabelPage.CheckTotalNumber() == 1)
            {
                trolleyDetailedLabelPage.ClearDownloads();
                PrintReportPage reportPage = trolleyDetailedLabelPage.PrintSimplified();
                var isReportGenerated = reportPage.IsReportGenerated();
                reportPage.Close();

                Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
            }
            else
            {
                Assert.Fail("Pas de Detailed Label Trolleys.");
            }
        }
    }
}