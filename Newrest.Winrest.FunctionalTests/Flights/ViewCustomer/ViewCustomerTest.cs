using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;

namespace Newrest.Winrest.FunctionalTests.Flights.ViewCustomer
{
    [TestClass]
    public class ViewCustomerTest : TestBase
    {
        // Impossible de évaluer l’expression, car un frame natif se trouve en haut de la pile des appels.
        private const int _timeout = 60 * 10 * 1000;
        private string serviceName1;
        private string serviceName2;
        private string site;
        private string loadingPlanName;
        private string customer;
        private string flightNumber;
        private DateTime serviceDateFrom;
        private DateTime serviceDateTo;
        private DateTime lpDateFrom;
        private DateTime lpDateTo;
        private string lpCartName;

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            serviceName1 = "service_1_" + DateUtils.Now.ToString("dd/MM/yyyy");
            serviceName2 = "service_2_" + DateUtils.Now.ToString("dd/MM/yyyy");
            site = "ACE";
            loadingPlanName = new Random().Next().ToString();
            customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            flightNumber = "flight_lp" + new Random().Next(1000, 5000).ToString();
            serviceDateFrom = DateUtils.Now;
            serviceDateTo = DateUtils.Now.AddMonths(2);
            lpDateFrom = serviceDateFrom.AddDays(2);
            lpDateTo = serviceDateTo.AddDays(-2);
            lpCartName = "lpcart bob" + new Random().Next().ToString();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(FL_FLIG_LoadingPlan_ServicesNotFound):
                    FL_FLIG_LoadingPlan_Initialize();
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
                case nameof(FL_FLIG_LoadingPlan_ServicesNotFound):
                    FL_FLIG_LoadingPlan_CleanUp();
                    break;

                default:
                    break;
            }
            base.TestCleanup();
        }

        private void FL_FLIG_LoadingPlan_Initialize()
        {
            // Prepare
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
            

            // LogIn
            var homePage = LogInAsAdmin();

            //CreateServices
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            if (servicePage.CheckTotalNumber() == 0)
            {
                var ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName1);
                var serviceGeneralInformationsPage = ServiceCreatePage.Create();
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, serviceDateFrom, serviceDateTo);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            }
            var isService1Creaated = servicePage.GetFirstServiceName().Contains(serviceName1);
            Assert.IsTrue(isService1Creaated, "Le service 1 n'a pas été créé.");

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName2);
            if (servicePage.CheckTotalNumber() == 0)
            {
                var ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName2);
                var serviceGeneralInformationsPage = ServiceCreatePage.Create();
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, serviceDateFrom, serviceDateTo);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName2);
            }
            var isService2Creaated = servicePage.GetFirstServiceName().Contains(serviceName2);
            Assert.IsTrue(isService2Creaated, "Le service 2 n'a pas été créé.");

            // Create LpCart
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            var lpCartModalCreate = lpCartPage.LpCartCreatePage();
            var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCartWithRoutes(code, lpCartName, site, customer, aircraft, lpDateFrom, lpDateTo, comment, route);
            LpCartDetailPage.BackToList();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
            var lpCartNumber = lpCartPage.CheckTotalNumber();
            Assert.AreNotEqual(0, lpCartNumber, "Le LpCart n'a pas été créé.");

            // Create LoadingPlan
            homePage.Navigate();
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
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.StartDate, lpDateFrom);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.EndDate, lpDateTo);
            var lpNumber = lpCartPage.CheckTotalNumber();
            Assert.AreNotEqual(0, lpNumber, "Le LoadingPlan n'a pas été créé.");

            //Create Flight
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, lpDateFrom);
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.SetDateState(lpDateFrom);
            var totalFlight = lpCartPage.CheckTotalNumber();
            Assert.AreNotEqual(0, totalFlight, "Le flight n'a pas été créé.");
        }

        private void FL_FLIG_LoadingPlan_CleanUp()
        {

            var homePage = LogInAsAdmin();
            //Delete Flight
            var flightPage = homePage.GoToFlights_FlightPage();
            var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
            massiveDeleteFlightPage.SetFlightName(flightNumber);
            massiveDeleteFlightPage.ClickSearchButton();
            massiveDeleteFlightPage.SelectFirstFlight();
            massiveDeleteFlightPage.Delete();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var result = flightPage.CheckTotalNumber();
            var isFlightDeleted = result == 0;
            Assert.IsTrue(isFlightDeleted, "flight not deleted");

            //Delete LoadingPlan
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, null, null, lpDateTo.AddMonths(2));
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            int totalNumber = loadingPlanPage.CheckTotalNumber();
            Assert.AreEqual(0, totalNumber, "La massive delete ne fonctionne pas.");

            //Delete LpCart
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
            lpCartPage.DeleteLpCart();
            var isLpCartDeleted = lpCartPage.CheckTotalNumber() == 0;
            Assert.IsTrue(isLpCartDeleted, "Le lpCart n'est pas supprimé");

            //Delete 2 services
            var servicePage = homePage.GoToCustomers_ServicePage();
            //Delete service1
            var serviceMassiveDelete = servicePage.ClickMassiveDelete();
            serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName1);
            serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
            serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy"));
            serviceMassiveDelete.ClickSearchButton();
            serviceMassiveDelete.DeleteFirstService();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            var total = servicePage.CheckTotalNumber();
            var isService1Deleted = total == 0;
            Assert.IsTrue(isService1Deleted, "La suppression d'un service ne fonctionne pas.");
            //Delete service2
            serviceMassiveDelete = servicePage.ClickMassiveDelete();
            serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName2);
            serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
            serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy"));
            serviceMassiveDelete.ClickSearchButton();
            serviceMassiveDelete.DeleteFirstService();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName2);
            total = servicePage.CheckTotalNumber();
            var isService2Deleted = total == 0;
            Assert.IsTrue(isService2Deleted, "La suppression d'un service ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_FLIG_LoadingPlan_ServicesNotFound()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.SetDateState(lpDateFrom);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            var servicesNames = editPage.GetListService();
            var servicesExist = servicesNames.Contains(serviceName1) && servicesNames.Contains(serviceName2);
            Assert.IsTrue(servicesExist, $"Les services rattachés au LP {loadingPlanName} ne s'affichent pas.");
        }
    }
}
