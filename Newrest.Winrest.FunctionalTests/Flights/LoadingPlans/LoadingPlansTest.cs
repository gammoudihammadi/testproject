using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.User;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;

namespace Newrest.Winrest.FunctionalTests.Flights
{
    [TestClass]
    public class LoadingPlansTest : TestBase
    {
        // Impossible de évaluer l’expression, car un frame natif se trouve en haut de la pile des appels.
        private const int _timeout = 60 * 10 * 1000;
        //________________________________ CREATE_LOADING PLAN _______________________________________
        //Data for testing
        private const string LP_NAME_CUSTOMER_INAC = "LPTestCustomInactive";
        private const string LP_NAME_CUSTOMER_INAC2 = "LPTestCustomInactive2";
        private const string LP_NAME_SITE_INAC = "LPTestSiteInactive";
        private const string LP_NAME_SITE_INAC2 = "LPTestSiteInactive2";
        private const string SITE = "ACE";
        private const string SITE_INACTIVE = "AEA";
        private const string CUSTOMER = "CAT Genérico";
        private const string CUSTOMER2 = "CANARYFLY";
        private const string ROUTE = "ACE-BCN";
        private const string AIRCRAFT = "ATX";
        string loadingPlanName = new Random().Next().ToString();

        // [TestMethod]
        public void Create_Lp_CardBob()
        {
            // Prepare
            string code = "Bob";
            string name = "lpcart bob";
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string aircraft = "AB320";
            DateTime from = DateUtils.Now.AddDays(5);
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


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Create_New_LoadingPlan()
        {
            // Prepare
            string aircraft = "AB318";
            string site = TestContext.Properties["SiteLP"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

                string nameLoadingPlan = loadingPlanName.ToString();
                string firstNameLoadingPlan = loadingPlanPage.GetFirstLoadingPlanName();

                //Assert
                Assert.AreEqual(nameLoadingPlan, firstNameLoadingPlan, "Aucun loading plan créé.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Create_New_LoadingPlanBob()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            string type = "BuyOnBoard";


            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site, type);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                loadingPlanPage = loadingPlanDetailsPage.BackToList();


                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

                string nameLoadingPlan = loadingPlanName.ToString();
                string firstNameLoadingPlan = loadingPlanPage.GetFirstLoadingPlanName();
                string firstNameLoadingPlanType = loadingPlanPage.GetFirstLoadingPlanNameType();
                //Assert
                Assert.AreEqual(nameLoadingPlan, firstNameLoadingPlan, "Aucun loading plan bob créé.");
                Assert.AreEqual(type, firstNameLoadingPlanType, "Aucun loading plan bob créé.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Create_New_LoadingPlanBobWithLpCart()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = "SMARTWINGS, A.S. (TVS)";
            string loadingPlanName = new Random().Next().ToString();
            string type = "BuyOnBoard";
            string guestName = "BOB";
            string serviceName = "Service for BOB";

            //Lp Cart
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string name = "lpcart bob" + DateUtils.Now.ToString("dd/MM/yyyy");
            string code = "Bob" + DateUtils.Now.ToString("dd/MM/yyyy");
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();
            string comment = "Bob comment";
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            // Arrange
            var homePage = LogInAsAdmin();

            //On verifie le guestType for return dans le customer
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customerLp);
            var customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            if (!customerGeneralInformationsPage.isBuyOnBoard())
            {
                customerGeneralInformationsPage.ActiveBuyOnBoard();
            }

            var customerBoBPage = customerGeneralInformationsPage.ClickBobTab();
            customerBoBPage.SelectGuestValueOnBob(guestName);
            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            ServiceGeneralInformationPage serviceGeneralInformationPage = servicePage.SelectFirstRow();
            ServicePricePage servicePricePage = serviceGeneralInformationPage.GoToPricePage();
            servicePricePage.UnfoldAll();
            ServiceCreatePriceModalPage servicePriceModalPage = servicePricePage.EditPrice(site, customerLp);
            var serviceDateFrom = servicePriceModalPage.GetServicePeriodFrom();
            var serviceDateTo = servicePriceModalPage.GetServicePeriodTo();
            servicePriceModalPage.Close();
            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();

            //verifie si le lp cart existe
            //Filter by Search
            lpCartPage.ResetFilter();
            // Create
            var lpCartModalCreate = lpCartPage.LpCartCreatePage();
            DateTime lpCartDateFrom = DateTime.ParseExact(serviceDateFrom, "dd/MM/yyyy", CultureInfo.InvariantCulture).AddDays(2);
            DateTime lpCartDateTo = DateTime.ParseExact(serviceDateTo, "dd/MM/yyyy", CultureInfo.InvariantCulture).AddDays(-2);
            var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, name, site, customer, aircraft, lpCartDateFrom, lpCartDateTo, comment);

            //Acces Onglet Cart                
            LpCartDetailPage.ClickAddtrolley();
            LpCartDetailPage.AddTrolley(trolleyName);

            //Create LpCart Scheme
            var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
            lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");
            LpCartDetailPage.BackToList();
            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create
            try
            {
                DateTime endDateLp = lpCartDateTo.AddDays(7);
                DateTime startDateLp = lpCartDateFrom.AddDays(-7);
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site, type);
                loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(endDateLp, startDateLp);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //add leg and service bob 
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);

                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

                //Edit the first loading plan
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                string lpCartCreated = code + " - " + name;
                generalInformationsPage.SelectLpCart(lpCartCreated);

                loadingPlanDetailsPage.ClickOnGeneralInformation();
                generalInformationsPage.ClickLoadingPlanLPCartEditLabels(trolleyScheme);


                string lpCartName = generalInformationsPage.GetLpCartName();
                //Assert
                Assert.AreEqual(code + " - " + name, lpCartName, "Aucun loading plan bob créé.");
                generalInformationsPage.SelectLpCart("None");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
                lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, name);
                lpCartPage.DeleteLpCart();
            }
        }

        //________________________________ FIN CREATE_LOADING PLAN ___________________________________

        //________________________________ FILTER LOADING PLAN _______________________________________



        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_Search()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName.ToString());

                Assert.AreEqual(loadingPlanName.ToString(), loadingPlanPage.GetFirstLoadingPlanName(), MessageErreur.FILTRE_ERRONE, "Search");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_SortBy()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site, null, "00", "10", "22", "00");
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                // Début Execution SortBy
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);

                if (!loadingPlanPage.isPageSizeEqualsTo100())
                {
                    loadingPlanPage.PageSize("8");
                    loadingPlanPage.PageSize("100");
                }

                loadingPlanPage.Filter(LoadingPlansPage.FilterType.SortBy, "Name");
                var isSortedByName = loadingPlanPage.IsSortedByName();

                loadingPlanPage.Filter(LoadingPlansPage.FilterType.SortBy, "Date");
                var isSortedByStartDate = loadingPlanPage.IsSortedByDate();

                loadingPlanPage.Filter(LoadingPlansPage.FilterType.SortBy, "ETD");
                var isSortedByETD = loadingPlanPage.IsSortedByEtd();

                // Assert    
                Assert.IsTrue(isSortedByName, MessageErreur.FILTRE_ERRONE, "Sort by name");
                Assert.IsTrue(isSortedByStartDate, MessageErreur.FILTRE_ERRONE, "Sort by date");
                Assert.IsTrue(isSortedByETD, MessageErreur.FILTRE_ERRONE, "Sort by ETD ");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_Site()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                // Create
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);

                if (!loadingPlanPage.isPageSizeEqualsTo100())
                {
                    loadingPlanPage.PageSize("8");
                    loadingPlanPage.PageSize("100");
                }

                // On vérifie que le LP créé sur le site recherché est présent dans la liste
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

                string nameLoadingPlan = loadingPlanName.ToString();
                string firstNameLoadingPlan = loadingPlanPage.GetFirstLoadingPlanName();

                Assert.AreEqual(nameLoadingPlan, firstNameLoadingPlan, MessageErreur.FILTRE_ERRONE, "Site");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_Airport()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                // Create
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);

                if (!loadingPlanPage.isPageSizeEqualsTo100())
                {
                    loadingPlanPage.PageSize("8");
                    loadingPlanPage.PageSize("100");
                }

                loadingPlanPage.Filter(LoadingPlansPage.FilterType.AirportDeparture, site);

                // On vérifie que le LP créé sur l'airport recherché est présent dans la liste
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);


                string nameLoadingPlan = loadingPlanName.ToString();
                string firstNameLoadingPlan = loadingPlanName.ToString();

                Assert.AreEqual(nameLoadingPlan, firstNameLoadingPlan, MessageErreur.FILTRE_ERRONE, "Airport Departure");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_Route()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                // Create
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);

                if (!loadingPlanPage.isPageSizeEqualsTo100())
                {
                    loadingPlanPage.PageSize("8");
                    loadingPlanPage.PageSize("100");
                }

                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Route, route);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

                Assert.AreEqual(loadingPlanName.ToString(), loadingPlanPage.GetFirstLoadingPlanName(), MessageErreur.FILTRE_ERRONE, "Route");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_ExpiredServices()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            string guestName = "YC";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            string customer = customerLp.Substring(0, customerLp.IndexOf("("));

            // Création d'un service pour le test
            CheckService(homePage, serviceName, site, customer, DateUtils.Now.AddYears(-1), DateUtils.Now.AddYears(+3));

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            loadingPlanPage.ResetFilter();

            // Create
            var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();

            try
            {
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtn();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);
                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                // Create
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);

                if (!loadingPlanPage.isPageSizeEqualsTo100())
                {
                    loadingPlanPage.PageSize("8");
                    loadingPlanPage.PageSize("100");
                }

                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.ExpiredService, true);
                Assert.AreEqual(loadingPlanPage.CheckTotalNumber(), 0, MessageErreur.FILTRE_ERRONE, "With expired services only");

                try
                {
                    CheckService(homePage, serviceName, site, customer, DateUtils.Now.AddMonths(-1), DateUtils.Now.AddDays(-5));

                    loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                    // Create
                    loadingPlanPage.ResetFilter();
                    loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                    loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
                    loadingPlanPage.Filter(LoadingPlansPage.FilterType.ExpiredService, true);

                    Assert.AreNotEqual(loadingPlanPage.CheckTotalNumber(), 0, MessageErreur.FILTRE_ERRONE, "With expired services only");
                }
                finally
                {
                    CheckService(homePage, serviceName, site, customer, DateUtils.Now.AddYears(-1), DateUtils.Now.AddYears(+3));
                }
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        private void CheckService(HomePage homePage, string serviceName, string site, string customer, DateTime dateFrom, DateTime dateTo)
        {
            var servicesPage = homePage.GoToCustomers_ServicePage();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicesPage.CheckTotalNumber() == 0)
            {
                var serviceModalPage = servicesPage.ServiceCreatePage();
                serviceModalPage.FillFields_CreateServiceModalPage(serviceName, "", "");
                var serviceGeneralInformationsPage = serviceModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddYears(-1), DateUtils.Now.AddYears(+3));
                servicesPage = pricePage.BackToList();

                servicesPage.Filter(ServicePage.FilterType.Search, serviceName);

                Assert.IsTrue(servicesPage.GetFirstServiceName().Contains(serviceName), MessageErreur.FILTRE_ERRONE, "Search");
            }
            else
            {
                var pricePage = servicesPage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(customer, site, dateFrom, dateTo);
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_Customers()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string customerLpFilter = TestContext.Properties["CustomerLpFilter"].ToString();
            string loadingPlanName = new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                //Assert
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);

                if (!loadingPlanPage.isPageSizeEqualsTo100())
                {
                    loadingPlanPage.PageSize("8");
                    loadingPlanPage.PageSize("100");
                }

                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Customers, customerLpFilter);

                bool verifCustomer = loadingPlanPage.VerifyCustomer(customerLp);
                Assert.IsTrue(verifCustomer, MessageErreur.FILTRE_ERRONE, "Customers");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        //________________________________ FIN FILTER LOADING PLAN _______________________________________

        //________________________________ UPDATE LOADING PLAN _______________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Update_General_Informations()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            var endDate = DateUtils.Now.AddMonths(3);


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

                // Edit the first loading plan
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                generalInformationsPage.EditLoadingPlanInformations(endDate);
                generalInformationsPage.BackToList();

                // Assert
                var newEndDate = loadingPlanPage.GetEndDate(decimalSeparatorValue);
                Assert.AreEqual(endDate.ToShortDateString(), newEndDate, "Aucune mise à jour n'a été effectuée.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Update_Details()
        {
            // Prepare          
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            string serviceName2 = TestContext.Properties["TrolleyService"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string guestName = "YC";
            string loadingPlanName = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                loadingPlanDetailsPage.ClickGuestBtn();
                loadingPlanDetailsPage.AddServiceBtn();
                var initValue = loadingPlanDetailsPage.GetServiceName();
                loadingPlanDetailsPage.AddNewService(serviceName2);
                loadingPlanDetailsPage.ClickOnDetailsPage();
                var newValue = loadingPlanDetailsPage.GetServiceName();

                // Assert
                Assert.AreNotEqual(initValue, newValue, "Le service n'a pas été mis à jour dans " + loadingPlanName);
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        //________________________________ FIN UPDATE LOADING PLAN _______________________________________

        //________________________________ CREATE GUEST _______________________________________________
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Create_New_Guest()
        {
            // Prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            string guestName = "YC";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                // Assert
                Assert.IsTrue(loadingPlanDetailsPage.IsGuestAdded(), "Aucun guest ajouté au loading plan.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        //________________________________ FIN CREATE GUEST ___________________________________________

        //________________________________ CREATE SERVICE _______________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Add_New_Service()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            string guestName = "YC";

            //Arrange
            var homePage = LogInAsAdmin();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create
            try
            {
                LoadingPlansCreateModalPage loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                LoadingPlansDetailsPage loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtn();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);
                //loadingPlanDetailsPage.ClickOnGeneralInformation();
                //loadingPlanDetailsPage.ClickOnDetailsPage();
                bool isServiceAdded = loadingPlanDetailsPage.IsServiceAdded(serviceName);
                // Assert
                Assert.IsTrue(isServiceAdded, "Le service {0} n'a pas été ajouté.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Add_Several_Services()
        {
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            string serviceBob = TestContext.Properties["ServiceBob"].ToString();
            string serviceTrolley = TestContext.Properties["TrolleyService"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            string guestName = "YC";
            HomePage homePage = LogInAsAdmin();
            LoadingPlansPage loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                LoadingPlansCreateModalPage loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                LoadingPlansDetailsPage loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                loadingPlanDetailsPage.ClickGuestBtn();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);
                loadingPlanDetailsPage.ClickOnDetailsPage();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceBob);
                loadingPlanDetailsPage.ClickOnDetailsPage();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceTrolley);
                loadingPlanDetailsPage.ClickOnDetailsPage();
                bool isServiceNameAdded = loadingPlanDetailsPage.IsServiceAdded(serviceName);
                bool isServiceBobAdded = loadingPlanDetailsPage.IsServiceAdded(serviceBob);
                bool isServiceTrolleyAdded = loadingPlanDetailsPage.IsServiceAdded(serviceTrolley);
                Assert.IsTrue(isServiceNameAdded, $"Le service {serviceName} n'a pas été ajouté.");
                Assert.IsTrue(isServiceBobAdded, $"Le service {serviceBob} n'a pas été ajouté.");
                Assert.IsTrue(isServiceTrolleyAdded, $"Le service {serviceTrolley} n'a pas été ajouté.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        //________________________________ FIN CREATE SERVICE ___________________________________________

        //________________________________ DELETE SERVICE _______________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Delete_Service()
        {
            // Prepare           
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string guestName = "YC";
            string loadingPlanName = new Random().Next(1000, 9000).ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtn();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);
                loadingPlanDetailsPage.ClickOnDetailsPage();
                loadingPlanDetailsPage.DeleteService();
                loadingPlanDetailsPage.ClickOnDetailsPage();
                // Assert
                Assert.IsTrue(loadingPlanDetailsPage.IsServiceDeleted(), "Le service n'a pas été supprimé.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        //________________________________ FIN DELETE SERVICE __________________________________________
        //________________________________ CHECK_FLIGHTS _______________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Check_Flights()
        {
            // Prepare            
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string siteTo = "AGP";
            string loadingPlanName = "000test-" + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNumber = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            try
            {
                // Act
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                // Create
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                loadingPlanCreateModalpage.Create();

                // Créer un flight et lui associer le loading plan
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightModalPage = flightPage.FlightCreatePage();
                flightModalPage.FillField_CreatNewFlight(flightNumber, customerLp, aircraft, site, siteTo, loadingPlanName);

                // Retourner dans loading plan, onglet Flight
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                var associatedFlightPage = generalInformationsPage.ClickOnFlightBtn();

                // On filtre pour une date de - 3 jours pour voir le vol créé
                associatedFlightPage.SetStartDate(DateUtils.Now.AddDays(-3));

                // Assert
                Assert.AreNotEqual(associatedFlightPage.GetFlightNumber(), "0", "Aucun numéro de vol associé au loading plan.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_Flights_Search()
        {
            // Prepare            
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string siteTo = "AGP";
            string loadingPlanName = "000test-" + DateUtils.Now.ToString("dd/MM/yyyy") + new Random().Next();
            string flightNumber = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            try
            {
                // Act
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                // Create
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                loadingPlanCreateModalpage.Create();

                // Créer un flight et lui associer le loading plan
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightModalPage = flightPage.FlightCreatePage();
                flightModalPage.FillField_CreatNewFlight(flightNumber, customerLp, aircraft, site, siteTo, loadingPlanName);

                // Retourner dans loading plan, onglet Flight
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName.ToString());
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                var associatedFlightPage = generalInformationsPage.ClickOnFlightBtn();

                // On filtre
                associatedFlightPage.ResetFilter();
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.Search, flightNumber);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.StartDate, DateUtils.Now);

                // Assert
                Assert.AreNotEqual(associatedFlightPage.GetFlightNumber(), "0", "Aucun numéro de vol associé au loading plan.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_Flights_ShowNotLinked()
        {
            // Prepare            
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string siteTo = "AGP";
            string loadingPlanName1 = "000test-" + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNumber = new Random().Next().ToString();
            string flightNumber1 = new Random().Next().ToString() + "not linked";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            try
            {
                // Act
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                // Create
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName1, customerLp, route, aircraft, site);
                loadingPlanCreateModalpage.Create();

                // Créer un flight et lui associer le loading plan
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightModalPage = flightPage.FlightCreatePage();
                flightModalPage.FillField_CreatNewFlight(flightNumber, customerLp, aircraft, site, siteTo, loadingPlanName1);

                var flightModalPage3 = flightPage.FlightCreatePage();
                flightModalPage3.FillField_CreatNewFlight(flightNumber1, customerLp, aircraft, site, siteTo);

                // Retourner dans loading plan, onglet Flight
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName1);
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                var associatedFlightPage = generalInformationsPage.ClickOnFlightBtn();

                // On filtre
                associatedFlightPage.ResetFilter();
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.StartDate, DateUtils.Now);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.EndDate, DateUtils.Now);


                // Assert
                Assert.AreNotEqual(associatedFlightPage.GetFlightNumber(), "0", "Aucun numéro de vol associé au loading plan.");

                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.StartDate, DateUtils.Now);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.EndDate, DateUtils.Now);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.Showflightsnotlinked, true);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.Search, flightNumber1);

                // Assert
                Assert.AreNotEqual(associatedFlightPage.GetFlightNumber(), "0", "Aucun numéro de vol associé au loading plan.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName1);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_Flights_Linked_LP_Deleted()
        {
            // Prepare            
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string siteTo = "AGP";
            string loadingPlanName = "000test-" + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNumber1 = new Random().Next().ToString() + " For linked";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            // Create
            var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
            loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
            loadingPlanCreateModalpage.Create();

            // Créer un flight et lui associer le loading plan
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, site);


            var flightModalPage3 = flightPage.FlightCreatePage();
            flightModalPage3.FillField_CreatNewFlight(flightNumber1, customerLp, aircraft, site, siteTo);

            // Retourner dans loading plan, onglet Flight
            loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
            var associatedFlightPage = generalInformationsPage.ClickOnFlightBtn();

            // On filtre
            associatedFlightPage.ResetFilter();

            associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.StartDate, DateUtils.Now.AddDays(1));
            associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.EndDate, DateUtils.Now.AddDays(1));
            associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.Showflightsnotlinked, true);
            associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.Search, flightNumber1);

            // Assert
            Assert.AreNotEqual(associatedFlightPage.GetFlightNumber(), "0", "Aucun numéro de vol associé au loading plan.");

            DeleteLoadingPlan(loadingPlanName);
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, flightNumber1);
            Assert.AreEqual(loadingPlanPage.CheckTotalNumber(), 0, "Loading plan n'est pas supprimer.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_Flights_Linked()
        {
            // Prepare            
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string siteTo = "AGP";
            string loadingPlanName = "000test-" + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNumber1 = new Random().Next().ToString() + " For linked";
            string loadingPlanName2 = "111test-" + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            try
            {
                // Act
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                // Create
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                loadingPlanCreateModalpage.Create();

                // Créer un flight et lui associer le loading plan
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);


                var flightModalPage3 = flightPage.FlightCreatePage();
                flightModalPage3.FillField_CreatNewFlight(flightNumber1, customerLp, aircraft, site, siteTo);

                // Retourner dans loading plan, onglet Flight
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName.ToString());
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                var associatedFlightPage = generalInformationsPage.ClickOnFlightBtn();

                // On filtre
                associatedFlightPage.ResetFilter();

                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.StartDate, DateUtils.Now.AddDays(1));
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.EndDate, DateUtils.Now.AddDays(1));
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.Showflightsnotlinked, true);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.Search, flightNumber1);


                // Assert
                Assert.AreNotEqual(associatedFlightPage.GetFlightNumber(), "0", "Aucun numéro de vol associé au loading plan.");

                associatedFlightPage.LinkNewFlight();
                associatedFlightPage.ResetFilter();

                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.StartDate, DateUtils.Now.AddDays(1));
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.EndDate, DateUtils.Now.AddDays(1));
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.Search, flightNumber1);

                // Assert
                Assert.AreNotEqual(associatedFlightPage.GetFlightNumber(), "0", "Aucun numéro de vol associé au loading plan.");

                //Même test avec 2 loading plan
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                //// Create un 2ème Loading plan
                loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName2, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                var loadingPlansGeneralInformationsPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                var loadingPlansFlightPage = loadingPlansGeneralInformationsPage.ClickOnFlightBtn();

                Assert.AreEqual(loadingPlansFlightPage.GetFlightNumber(), "1", "Le vol numéro {0} de vol n'est pas visible sur Show All du 2ème loading plan crée {1}", flightNumber1, loadingPlanName2);

                loadingPlansFlightPage.Filter(LoadingPlansFlightPage.FilterType.ShowLinkedFlights, true);

                //// Assert
                Assert.AreEqual(loadingPlansFlightPage.GetFlightNumber(), "0", "Le vol numéro {0} est déjà associé", flightNumber1);

                loadingPlansFlightPage.Filter(LoadingPlansFlightPage.FilterType.Showflightsnotlinked, true);
                Assert.AreEqual(loadingPlansFlightPage.GetFlightNumber(), "1", "Le vol numéro {0} de vol n'est pas visible sur Show flights not linked", flightNumber1);

                loadingPlansFlightPage.LinkNewFlight();

                // Assert
                Assert.AreEqual(associatedFlightPage.GetFlightNumber(), "0", "Aucun numéro de vol associé au loading plan.");

                loadingPlansFlightPage.Filter(LoadingPlansFlightPage.FilterType.ShowLinkedFlights, true);
                // Assert
                Assert.AreEqual(associatedFlightPage.GetFlightNumber(), "1", "Aucun numéro de vol associé au loading plan.");

            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
                DeleteLoadingPlan(loadingPlanName2);
            }
        }

        private void DeleteLoadingPlan(string loadingPlanName)
        {
            var winrestUrl = TestContext.Properties["Winrest_URL"].ToString();

            WebDriver.Navigate().GoToUrl(winrestUrl);

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Suppression du loading plan créé
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName.ToString());

            if (loadingPlanPage.CheckTotalNumber() > 0)
            {
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                generalInformationsPage.DeleteLoadingPlan();
            }
        }
        private void DeleteService(string serviceNameToDelete)
        {
            var winrestUrl = TestContext.Properties["Winrest_URL"].ToString();

            WebDriver.Navigate().GoToUrl(winrestUrl);

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Suppression du service
            var servicePage = homePage.GoToCustomers_ServicePage();
            ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
            serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceNameToDelete);
            serviceMassiveDelete.ClickSearchButton();
            serviceMassiveDelete.DeleteFirstService();
        }

        //________________________________ FIN CHECK_FLIGHTS ___________________________________________

        //________________________________ CHECK_LINKED LP ___________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Check_Linked_Loading_Plan()
        {
            // Prepare           
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            string loadingPlanNameSplit = String.Format("{0}_split", loadingPlanName.ToString());

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName.ToString());

                // Create
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();

                var loadingPlanSplitModalPage = generalInformationsPage.SplitLoadingPlanPage();
                generalInformationsPage = loadingPlanSplitModalPage.FillFields_SplitLoadingPlan(loadingPlanNameSplit);

                var linkedLoadingPlansPage = generalInformationsPage.RedirectToLinkedLoadingPlansPage();
                var newValue = linkedLoadingPlansPage.GetLinkedLoadingPlansNumber();

                // Assert 
                Assert.AreNotEqual("0", newValue, "Aucun loading plan associé au loading plan en cours.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        //________________________________ FIN CHECK_LINKED LP ___________________________________________

        //________________________________ DUPLICATE LP __________________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Duplicate_Loading_Plan()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            var startDate = DateUtils.Now.AddDays(1);
            var endDate = DateUtils.Now.AddMonths(3);
            string loadingPlanNameBis = String.Format("{0}_bis", loadingPlanName.ToString());

            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanPage = loadingPlanDetailsPage.BackToList();
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName.ToString());

                // Récupérer le loading plan créé
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();

                // Create
                var loadingPlanDuplicationModalPage = generalInformationsPage.DuplicateLoadingPlanPage();
                loadingPlanDuplicationModalPage.FillFields_DuplicateLoadingPlan(loadingPlanNameBis, startDate, endDate);
                generalInformationsPage.BackToList();

                //Assert
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);
                string firstLoadingPlanName = loadingPlanPage.GetFirstLoadingPlanName();
                Assert.AreEqual(loadingPlanNameBis, firstLoadingPlanName, "Le loading plan n'a pas été dupliqué.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanNameBis);
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        //________________________________ FIN DUPLICATE LP ______________________________________________

        //________________________________ SPLIT LP ______________________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Split_Loading_Plan()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            string loadingPlanNameSplit = String.Format("{0}_split", loadingPlanName.ToString());

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName.ToString());

                // Récupérer le loading plan créé
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();

                // Create
                var linkedLoadingPlansPage = generalInformationsPage.RedirectToLinkedLoadingPlansPage();
                var initValue = linkedLoadingPlansPage.GetLinkedLoadingPlansNumber();

                linkedLoadingPlansPage.ClickOnGeneralInformation();
                var loadingPlanSplitModalPage = generalInformationsPage.SplitLoadingPlanPage();
                generalInformationsPage = loadingPlanSplitModalPage.FillFields_SplitLoadingPlan(loadingPlanNameSplit);

                linkedLoadingPlansPage = generalInformationsPage.RedirectToLinkedLoadingPlansPage();
                var newValue = linkedLoadingPlansPage.GetLinkedLoadingPlansNumber();

                // Assert 
                Assert.AreNotEqual(initValue, newValue, "Le loading plan n'a pas été divisé.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete()
        {
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = "CAT Genérico";
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            var endDate = DateUtils.Now;
            HomePage homePage = LogInAsAdmin();
            LoadingPlansPage loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            LoadingPlansCreateModalPage loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
            loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
            LoadingPlansDetailsPage loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
            loadingPlanPage = loadingPlanDetailsPage.BackToList();
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            LoadingPlansGeneralInformationsPage generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
            generalInformationsPage.EditLoadingPlanInformations(endDate);
            generalInformationsPage.BackToList();
            loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, site, customerLp, endDate.AddDays(1));
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName.ToString());
            int totalNumber = loadingPlanPage.CheckTotalNumber();
            Assert.AreEqual(0, totalNumber, "La massive delete ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_ShowLinked()
        {
            // Prepare            
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string siteTo = "AGP";
            string loadingPlanName1 = "001test-" + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNumber = new Random().Next().ToString();


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            try
            {
                // Act
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                // Create
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName1, customerLp, route, aircraft, site);
                loadingPlanCreateModalpage.Create();

                // Créer un flight et lui associer le loading plan
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightModalPage = flightPage.FlightCreatePage();
                flightModalPage.FillField_CreatNewFlight(flightNumber, customerLp, aircraft, site, siteTo, loadingPlanName1);

                // Retourner dans loading plan, onglet Flight
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName1);
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                var associatedFlightPage = generalInformationsPage.ClickOnFlightBtn();

                // On filtre
                associatedFlightPage.ResetFilter();
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.StartDate, DateUtils.Now);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.EndDate, DateUtils.Now);

                // Assert
                Assert.AreNotEqual(associatedFlightPage.GetFlightNumber(), "0", "Aucun numéro de vol associé au loading plan.");

                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.StartDate, DateUtils.Now);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.EndDate, DateUtils.Now);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.ShowLinkedFlights, true);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.Search, flightNumber);

                // Assert
                Assert.AreEqual(associatedFlightPage.GetFlightNumber(), "1", "le numéro de vol est associé au loading plan.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName1);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_Showall()
        {
            // Prepare            
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string siteTo = "AGP";
            string loadingPlanName1 = "002test-" + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNumber = new Random().Next().ToString();
            string flightNumber1 = new Random().Next().ToString() + "not linked";

            //Arrange
            var homePage = LogInAsAdmin();

            try
            {
                // Act
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                // Create
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName1, customerLp, route, aircraft, site);
                loadingPlanCreateModalpage.Create();

                // Créer un flight et lui associer le loading plan
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightModalPage = flightPage.FlightCreatePage();
                flightModalPage.FillField_CreatNewFlight(flightNumber, customerLp, aircraft, site, siteTo, loadingPlanName1);

                var flightModalPage3 = flightPage.FlightCreatePage();
                flightModalPage3.FillField_CreatNewFlight(flightNumber1, customerLp, aircraft, site, siteTo);

                // Retourner dans loading plan, onglet Flight
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName1);
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                var associatedFlightPage = generalInformationsPage.ClickOnFlightBtn();

                // On filtre
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.StartDate, DateUtils.Now);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.EndDate, DateUtils.Now);
                associatedFlightPage.Filter(LoadingPlansFlightPage.FilterType.ShowAll, true);

                var checkFlightNumber = associatedFlightPage.GetFlightNumber();
                // Assert
                Assert.AreNotEqual(checkFlightNumber, "0", "les numéros du vol sont associé au loading plan.");
            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName1);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_History()
        {

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            loadingPlanPage.ResetFilter();
            var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
            var loadingPlanHistoryModalPage = generalInformationsPage.CheckLoadingPlanHistory();

            Assert.IsTrue(loadingPlanHistoryModalPage.verifyHistoryModal(), "le loading plan a un historique");


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Propagate()
        {

            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();


            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                var loadingPlanGeneralPage = loadingPlanDetailsPage.ClickOnGeneralInformation();
                LoadingPlansPropagateModalPage loadingPlanPropagatePage = loadingPlanGeneralPage.PropagateLoadingPlan();
                var PropagateSite = loadingPlanPropagatePage.SecurityCheckLoadingPlan();
                loadingPlanPage = loadingPlanDetailsPage.BackToList();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, PropagateSite);
                if (!loadingPlanPage.isPageSizeEqualsTo100())
                {
                    loadingPlanPage.PageSize("8");
                    loadingPlanPage.PageSize("100");
                }
                var exist = loadingPlanPage.VerifyLoadingPlan(loadingPlanName);
                Assert.IsTrue(exist, "la propagation du site sur loading plan ne fonction pas");

            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_customersearch()
        {
            //Prepare
            string customerName = "SMARTWINGS, A.S. (TVS)";
            //Arrange 
            HomePage homePage = LogInAsAdmin();
            //Act
            LoadingPlansPage loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            LoadingPlanMassiveDeleteModalPage loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            loadingPlanMassiveDeleteModalPage.Filter(LoadingPlanMassiveDeleteModalPage.FilterType.Customer, customerName);
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            bool result = loadingPlanMassiveDeleteModalPage.VerifyAllCustomers(customerName);
            //Arrange 
            Assert.IsTrue(result, "erreur de filtrage par customer");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_filter_fromto()
        {
            var dateFrom = DateTime.Now;
            var dateTo = DateTime.Now.AddDays(1);
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.StartDate, dateFrom);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.EndDate, dateTo);
            var result = loadingPlanPage.VerifyFromTo(dateFrom, dateTo);
            Assert.IsTrue(result, "erreur de filtrage par date dates out of range");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_sitesearch()
        {
            // Prepare            
            string site = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            loadingPlanMassiveDeleteModalPage.Filter(LoadingPlanMassiveDeleteModalPage.FilterType.Site, site);
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            loadingPlanMassiveDeleteModalPage.PageSize("100");
            var result = loadingPlanMassiveDeleteModalPage.VerifyAllSites(site);
            Assert.IsTrue(result, "erreur de filtrage par site");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_organizebyname()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            loadingPlanMassiveDeleteModalPage.PageSize("100");
            // Tri par LodingPlanName dans le premier sens, ascendant .
            var result = loadingPlanMassiveDeleteModalPage.VerifySortByLPNameASC();
            Assert.IsTrue(result, "Erreur de tri par LodingPlanName.");

            //Tri par LodingPlanName dans le deuxième sens ,descendant
            loadingPlanMassiveDeleteModalPage.SortByLoadingPlanNameDESC();
            result = loadingPlanMassiveDeleteModalPage.VerifySortByLPNameDESC();
            Assert.IsTrue(result, "Erreur de tri par LodingPlanName.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_inactivecustomersearch()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            var resultBeforeCheck = loadingPlanMassiveDeleteModalPage.VerifyCustomersHasInactive();
            Assert.IsFalse(resultBeforeCheck, "customer inactive appear even though show inactive customers is not checked");
            loadingPlanMassiveDeleteModalPage.ShowInactiveCustomers(true);
            var resutAfterCheck = loadingPlanMassiveDeleteModalPage.VerifyCustomersHasInactive();
            Assert.IsTrue(resutAfterCheck, "customer inactive not appearing even though show inactive customers is checked");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_organizebyservice()
        {
            HomePage homePage = LogInAsAdmin();
            LoadingPlansPage loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            LoadingPlanMassiveDeleteModalPage loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            loadingPlanMassiveDeleteModalPage.PageSize("100");
            loadingPlanMassiveDeleteModalPage.SortByServiceNumberASC();
            bool result = loadingPlanMassiveDeleteModalPage.VerifySortByServiceNumberASC();
            Assert.IsTrue(result, "Erreur de tri par Service Number.");
            loadingPlanMassiveDeleteModalPage.SortByServiceNumberDESC();
            result = loadingPlanMassiveDeleteModalPage.VerifySortByServiceNumberDESC();
            Assert.IsTrue(result, "Erreur de tri par Service Number.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_inactivecustomercheck()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            loadingPlanMassiveDeleteModalPage.ShowInactiveCustomers(true);
            loadingPlanMassiveDeleteModalPage.SelectInactiveCustomers();
            loadingPlanMassiveDeleteModalPage.Filter(LoadingPlanMassiveDeleteModalPage.FilterType.From, "08/03/2019");
            loadingPlanMassiveDeleteModalPage.Filter(LoadingPlanMassiveDeleteModalPage.FilterType.To, "27/04/2024");
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            var result = loadingPlanMassiveDeleteModalPage.VerifyAllCustomersAreInactive();
            Assert.IsTrue(result, "customers are not inactive");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_organizebysite()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            //loadingPlanMassiveDeleteModalPage.Filter(LoadingPlanMassiveDeleteModalPage.FilterType.Site, "AGP");
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            loadingPlanMassiveDeleteModalPage.PageSize("100");
            // Tri par Site dans le premier sens, ascendant .
            loadingPlanMassiveDeleteModalPage.SortBySiteASC();
            var result = loadingPlanMassiveDeleteModalPage.VerifySortBySiteASC();
            Assert.IsTrue(result, "Erreur de tri par Site.");

            //Tri par Site dans le deuxième sens ,descendant
            loadingPlanMassiveDeleteModalPage.SortBySiteDESC();
            result = loadingPlanMassiveDeleteModalPage.VerifySortBySiteDESC();
            Assert.IsTrue(result, "Erreur de tri par Site.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_inactivesitescheck()
        {
            HomePage homePage = LogInAsAdmin();
            LoadingPlansPage loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            LoadingPlanMassiveDeleteModalPage loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            bool resultBeforeCheck = loadingPlanMassiveDeleteModalPage.VerifySitesHasInactive();
            Assert.IsFalse(resultBeforeCheck, "site inactive appear even though show inactive sites is not checked");
            loadingPlanMassiveDeleteModalPage.ShowInactiveSites(true);
            bool resultAfterCheck = loadingPlanMassiveDeleteModalPage.VerifySitesHasInactive();
            Assert.IsTrue(resultAfterCheck, "site inactive not appearing even though show inactive sites is checked");
        }

        /**
         * SQL pour purge : delete from [ServiceToCustomerToSites] where SiteId in (select Id from Sites where Name ='AEA')
         */

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_status_inactivesites()
        {
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            string guestName = "YC";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //set site active for LP creation - AEA should be active for LP creation
            //Add site to user
            ActivateSiteForUser(homePage, SITE_INACTIVE);
            ActivateFirstSiteInList(homePage, SITE_INACTIVE);
            //Activate ACE
            ActivateSiteForUser(homePage, SITE);

            //Add a price for a service for the inactive site
            var servicePage = homePage.GoToCustomers_ServicePage();
            var pricePage = servicePage.ClickOnFirstService();
            //Adding price if it does not exist
            pricePage.SearchPriceForCustomer(CUSTOMER, SITE_INACTIVE, DateTime.Now.AddDays(-10), DateTime.Now.AddYears(1));
            string lpNameSiteInactive = LP_NAME_SITE_INAC + SITE_INACTIVE;
            string lpnameSiteActive = LP_NAME_SITE_INAC2 + SITE;

            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            Thread.Sleep(2000);
            CreateLoadingPlanForMassiveDeletePopupTests(loadingPlanPage, SITE_INACTIVE, CUSTOMER, lpNameSiteInactive, guestName, serviceName);
            Thread.Sleep(2000);
            CreateLoadingPlanForMassiveDeletePopupTests(loadingPlanPage, SITE, CUSTOMER, lpnameSiteActive, guestName, serviceName);
            Thread.Sleep(2000);

            //Deactivating for testing purpose 
            DeactivateFirstSiteInList(homePage, SITE_INACTIVE);
            //When you deactivate a site, you need to reaffect it
            ActivateSiteForUser(homePage, SITE_INACTIVE);
            ClearCache();

            loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();

            try
            {
                //select only inactive site 
                loadingPlanMassiveDeleteModalPage.SelectUseCaseByValue("Only Inactive Sites");
                loadingPlanMassiveDeleteModalPage.SetSearch(LP_NAME_SITE_INAC);
                loadingPlanMassiveDeleteModalPage.SetDateTimeTo(DateTime.Now.AddYears(1));
                loadingPlanMassiveDeleteModalPage.ClickSearchButton();

                //Assert
                Assert.IsTrue(loadingPlanMassiveDeleteModalPage.VerifyAllSitesAreInactive(), "Anomalie, presence site(s) active(s)");
            }
            finally
            {
                //Reset site active
                ActivateFirstSiteInList(homePage, SITE_INACTIVE);
                //Waiting for site to be set active before closing program
                Thread.Sleep(1500);
            }
        }

        private void ActivateSiteForUser(HomePage homePage, string site)
        {
            var userPage = homePage.GoToParameters_User();

            var adminName = TestContext.Properties["Admin_UserName"].ToString();
            string userName = adminName.Substring(0, adminName.IndexOf("@"));
            userPage.SearchAndSelectUserByClickingFirstRow(userName);
            userPage.ClickOnAffectedSite();
            userPage.ActivateSite(site);
        }

        private void CreateLoadingPlanForMassiveDeletePopupTests(LoadingPlansPage loadingPlanPage, string site, string customer, string lpName, string guestName, string serviceName)
        {
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, lpName);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.StartDate, DateUtils.Now.AddDays(-80));
            if (loadingPlanPage.CheckTotalNumber() == 0)
            {
                var createModal = loadingPlanPage.LoadingPlansCreatePage();
                createModal.FillField_CreateNewLoadingPlan(lpName, customer, ROUTE, AIRCRAFT, site);
                LoadingPlansDetailsPage loadingPlanDetailsPage = createModal.Create();
                // sans service le lp ne remonte pas dans l'index lp
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtn();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);

                loadingPlanDetailsPage.BackToList();
            }
            else
            {
                LoadingPlansGeneralInformationsPage lpPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                lpPage.EditLoadingPlanInformations(DateUtils.Now.AddDays(80));
                lpPage.BackToList();
            }
        }

        private void ActivateFirstSiteInList(HomePage homePage, string site)
        {
            ParametersSites sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, string.IsNullOrEmpty(site) ? SITE_INACTIVE : site);
            sitePage.ActivateFirstSiteInList();
        }

        private void DeactivateFirstSiteInList(HomePage homePage, string site)
        {
            ParametersSites sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, string.IsNullOrEmpty(site) ? SITE_INACTIVE : site);
            sitePage.DeactivateFirstSiteInList();
        }

        private void ChangeCustomerActiveStatus(bool setActive, HomePage homePage, string customer = null)
        {
            var customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.Filter(CustomerPage.FilterType.ShowAll, true);
            customerPage.Filter(CustomerPage.FilterType.Search, string.IsNullOrEmpty(customer) ? CUSTOMER2 : customer);
            var custInfoPage = customerPage.SelectFirstCustomer();

            if (setActive)
                custInfoPage.SetCustomerActive();
            else custInfoPage.SetCustomerInactive();
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_status_inactivecustomers()
        {
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            string guestName = "YC";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            CreateLoadingPlanForMassiveDeletePopupTests(loadingPlanPage, SITE, CUSTOMER, LP_NAME_CUSTOMER_INAC, guestName, serviceName);
            CreateLoadingPlanForMassiveDeletePopupTests(loadingPlanPage, SITE, CUSTOMER2, LP_NAME_CUSTOMER_INAC2, guestName, serviceName);

            try
            {
                //set customer inactive
                ChangeCustomerActiveStatus(false, homePage);

                //select only inactive site
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
                loadingPlanMassiveDeleteModalPage.SelectUseCaseByValue("Only Inactive Customers");
                loadingPlanMassiveDeleteModalPage.SetSearch(LP_NAME_CUSTOMER_INAC);
                loadingPlanMassiveDeleteModalPage.SetDateTimeTo(DateTime.Now.AddYears(1));
                loadingPlanMassiveDeleteModalPage.ClickSearchButton();

                //Assert
                Assert.IsTrue(loadingPlanMassiveDeleteModalPage.VerifyAllCustomersAreInactive(), "Error : found customer(s) active(s)");
            }
            finally
            {
                //Reset customer active
                ChangeCustomerActiveStatus(true, homePage);
                //Waiting for customer to be set active before closing program
                Thread.Sleep(1500);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_status_activesite()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            Thread.Sleep(4000);
            //select only inactive site 
            loadingPlanMassiveDeleteModalPage.SelectActiveSites();
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            loadingPlanMassiveDeleteModalPage.PageSize("100");
            var queDesSitesActives = loadingPlanMassiveDeleteModalPage.CheckIfSitesAreActive();

            //Assert
            Assert.IsTrue(queDesSitesActives, "Anomalie, presence site(s) inactive(s)");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_link()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();

            var loadingPlanDetailPage = loadingPlanMassiveDeleteModalPage.ClickFirstLink();
            var result = loadingPlanDetailPage.isLoadingPlanDetailPage();
            Assert.IsTrue(result, "Loading Plan Details page not showing");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_namesearch()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            var firstLoadingPlanName = loadingPlanMassiveDeleteModalPage.getLodingPlanNames().First();
            loadingPlanMassiveDeleteModalPage.Filter(LoadingPlanMassiveDeleteModalPage.FilterType.loadingPlanName, firstLoadingPlanName);
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            Assert.IsTrue(loadingPlanMassiveDeleteModalPage.VerifySearchResult(firstLoadingPlanName));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_selectall()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            loadingPlanMassiveDeleteModalPage.PageSize("100");
            int pageNumber = loadingPlanMassiveDeleteModalPage.getNumberPage();
            if (pageNumber < 2)
            {
                loadingPlanMassiveDeleteModalPage.Filter(LoadingPlanMassiveDeleteModalPage.FilterType.From, DateUtils.Now.AddMonths(-3).ToString("dd/MM/yyyy"));
                loadingPlanMassiveDeleteModalPage.Filter(LoadingPlanMassiveDeleteModalPage.FilterType.To, DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy"));
                loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            }
            var totalNumber = loadingPlanMassiveDeleteModalPage.CheckTotalNumber();
            var result = loadingPlanMassiveDeleteModalPage.VerifySelectedAll(totalNumber);
            //Assert
            Assert.IsTrue(result, "Erreur de Select All,Le nombre de lignes cochées et l'affichage du nombre de sélections ne sont pas égaux.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_pagination()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            var totalNumber = loadingPlanMassiveDeleteModalPage.CheckTotalNumber();
            //Assert
            Assert.IsTrue(loadingPlanMassiveDeleteModalPage.VerifyPagination(8, totalNumber), "Erreur de Pagination.");
            Assert.IsTrue(loadingPlanMassiveDeleteModalPage.VerifyPagination(50, totalNumber), "Erreur de Pagination.");
            Assert.IsTrue(loadingPlanMassiveDeleteModalPage.VerifyPagination(100, totalNumber), "Erreur de Pagination.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_status_all()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            Thread.Sleep(4000);

            loadingPlanMassiveDeleteModalPage.SelecAllSites();
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();
            loadingPlanMassiveDeleteModalPage.PageSize("100");
            var allSites = loadingPlanMassiveDeleteModalPage.CheckIfAllSitesArePresent();
            //Assert
            Assert.IsTrue(allSites);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_organizebycustomer()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            var loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
            loadingPlanMassiveDeleteModalPage.ClickSearchButton();

            loadingPlanMassiveDeleteModalPage.SortByCustomerDESC();
            var result = loadingPlanMassiveDeleteModalPage.VerifySortByCustomerDESC();
            Assert.IsTrue(result, "erreur de tri by customer desc");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Massivedelete_inactivesitessearch()
        {
            string site = "ACE";
            string inactive_site = "Inactive - ACE";
            HomePage homePage = LogInAsAdmin();
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
                LoadingPlansPage loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                LoadingPlanMassiveDeleteModalPage loadingPlanMassiveDeleteModalPage = loadingPlanPage.ClickMassiveDelete();
                loadingPlanMassiveDeleteModalPage.ShowInactiveSites(true);
                loadingPlanMassiveDeleteModalPage.SelectSites(inactive_site);
                loadingPlanMassiveDeleteModalPage.ClickSearchButton();
                int totalpage = 0;
                var isData = loadingPlanMassiveDeleteModalPage.CheckTotalNumber();
                if (isData > 0)
                {
                    loadingPlanMassiveDeleteModalPage.PageSize("100");
                    totalpage = loadingPlanMassiveDeleteModalPage.getNumberPage();
                }
                if (totalpage > 1)
                {
                    Assert.IsTrue(loadingPlanMassiveDeleteModalPage.VerifySiteSearch(site), "erreur de filtrage par site");
                    loadingPlanMassiveDeleteModalPage.goToPage(totalpage);
                    Assert.IsTrue(loadingPlanMassiveDeleteModalPage.VerifySiteSearch(site), "erreur de filtrage par site");
                }
                else Assert.IsTrue(loadingPlanMassiveDeleteModalPage.VerifySiteSearch(site), "erreur de filtrage par site");
            }
            finally
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
        }



        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_selectService_Update_Flight()
        {
            // Prepare
            string aircraft = "AB318";
            string site = TestContext.Properties["SiteLP"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraftBob = "BCS1";
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
            loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
            var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
            loadingPlanPage = loadingPlanDetailsPage.BackToList();
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            loadingPlanPage.ClickOnFirstLoadingPlan();
            loadingPlanDetailsPage.ClickOnDetailsPage();
            loadingPlanDetailsPage.ClickAddGuestBtn();
            loadingPlanDetailsPage.ClickCreateGuestBtn();
            loadingPlanDetailsPage.ClickFirstGuest();
            loadingPlanDetailsPage.AddServiceBtn();
            loadingPlanDetailsPage.AddNewService("FirstService");
            loadingPlanPage = loadingPlanDetailsPage.BackToList();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerLp, aircraft, siteFrom, siteTo, loadingPlanName);
            var loadingsPlanPage = homePage.GoToFlights_LoadingPlansPage();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            loadingPlanPage.ClickOnFirstLoadingPlan();
            loadingPlanDetailsPage.ClickOnDetailsPage();

            loadingPlanDetailsPage.ClickAddGuestBtn();
            loadingPlanDetailsPage.ClickCreateGuestBtn();
            loadingPlanDetailsPage.ClickFirstGuest();
            loadingPlanDetailsPage.AddServiceBtn();
            loadingPlanDetailsPage.AddNewService("Trolley Service");
            var flightPages = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, loadingPlanName);
            flightPage.ConsultFirstFlight();
            int NumberOfLines = flightPage.CheckTotalNumberOfLines();

            Assert.AreEqual(NumberOfLines, 1, "Update Fonctionne");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_UseCase_LPServiceMethod_MAD()
        {
            string aircraft = "AB318";
            string serviceForBOB = TestContext.Properties["ServiceBob"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();

            var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
            loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
            var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
            loadingPlanPage = loadingPlanDetailsPage.BackToList();
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            loadingPlanPage.ClickOnFirstLoadingPlan();
            loadingPlanDetailsPage.ClickOnDetailsPage();
            loadingPlanDetailsPage.ClickAddGuestBtn();
            loadingPlanDetailsPage.ClickCreateGuestBtn();
            loadingPlanDetailsPage.ClickFirstGuest();
            loadingPlanDetailsPage.AddServiceBtn();
            loadingPlanDetailsPage.AddNewServiceWithoutsave("AIR");
            loadingPlanDetailsPage.AddServiceBtn();
            loadingPlanDetailsPage.AddNewService(serviceForBOB);
            loadingPlanDetailsPage.ClickOnDetailsPage();
            var firstline = loadingPlanDetailsPage.GetFirstLineType();
            var secondline = loadingPlanDetailsPage.GetSecondLineType();

            Assert.IsTrue(firstline != null && secondline != null && firstline != secondline);

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_selectService()
        {
            //FIXME si je clique sur demain dans Flight, le vol apparait
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string route = "ACE - BCN";
            string customerLp = TestContext.Properties["CustomerNamebis"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            string guestName = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string newServiceName = "Service2- " + TestContext.Properties["ServiceNameLP"].ToString();
            string serviceName = "Service1-" + TestContext.Properties["ServiceNameLP"].ToString();
            string siteFrom = TestContext.Properties["SiteACE"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string flightNumber = "ACE" + new Random().Next().ToString();
            // Prepare Service
            string serviceCategory = TestContext.Properties["ServiceCategory"].ToString();
            string customerName = TestContext.Properties["CustomerLpFilter"].ToString();
            DateTime fromDate = DateUtils.Now.AddMonths(-1);
            DateTime toDate = DateUtils.Now.AddMonths(+2);
            DateTime startFlightDate = DateUtils.Now.AddDays(+1);
            DateTime endFlightDate = DateUtils.Now.AddDays(+7);
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            // create service 1
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            if (servicePage.CheckTotalNumber() == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage servicePricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage createpriceModalPage = servicePricePage.AddNewCustomerPrice();
                servicePricePage = createpriceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate);
                servicePage = servicePricePage.BackToList();
            }
            else
            {
                ServicePricePage servicePrice = servicePage.ClickOnFirstService();
                servicePrice.ToggleFirstPrice();
                ServiceCreatePriceModalPage servicePriceUpdateModalPage = servicePrice.EditFirstPrice(site, customerName);
                servicePriceUpdateModalPage.EditPriceDates(fromDate, toDate);
                servicePrice.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, newServiceName);
            if (servicePage.CheckTotalNumber() == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(newServiceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage servicePricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage createpriceModalPage = servicePricePage.AddNewCustomerPrice();
                servicePricePage = createpriceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate);
                servicePage = servicePricePage.BackToList();
            }
            else
            {
                ServicePricePage servicePrice = servicePage.ClickOnFirstService();
                servicePrice.ToggleFirstPrice();
                ServiceCreatePriceModalPage servicePriceUpdateModalPage = servicePrice.EditFirstPrice(site, customerName);
                servicePriceUpdateModalPage.EditPriceDates(fromDate, toDate);
                servicePrice.BackToList();
            }
            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                //create a loading Plan
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(toDate, fromDate);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //add leg and  two services
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);
                loadingPlanPage = loadingPlanDetailsPage.BackToList();

                // Go to flight page and add flight 
                var flightPage = homePage.GoToFlights_FlightPage();
                var flightMultiplModalpage = flightPage.CreateMultiFlights();
                flightMultiplModalpage.FillField_CreatMultiFlight(siteFrom, flightNumber, siteTo, aircraft, startFlightDate, endFlightDate, customerLp);

                // Go to LoadingPlan List and select the created Loading plan
                flightMultiplModalpage.CheckLoadingPlan(loadingPlanName);
                flightPage = flightMultiplModalpage.Validate();
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);

                //process of changing service
                var loadingPlanPageDetails = loadingPlanPage.ClickOnFirstLoadingPlan();
                loadingPlanPageDetails.GoToLoadingPlanDetailPage();
                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(newServiceName);
                bool servicesisUpdated = loadingPlanDetailsPage.GetServicesNames(serviceName, newServiceName);

                // Go to flight page and test the created Flight 
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                FlightEditModalPage flightEditPage = flightPage.EditFlightRow(flightNumber);
                FlightCreateModalPage editDetailsFlight = flightEditPage.OpenEditDetailsflightModal();
                editDetailsFlight.CheckLoadingPlan(loadingPlanName);
                editDetailsFlight.SaveFlightAndReloadLP();
                bool servicesisUpdatedFlight = flightEditPage.GetServicesNames(newServiceName);
                var y = servicesisUpdated && servicesisUpdatedFlight;

                //Assert
                Assert.IsTrue(servicesisUpdated && servicesisUpdatedFlight, "la mis a jour de service n'est pas appraitre dans tous les flights");
            }
            finally
            {

                DeleteLoadingPlan(loadingPlanName);
                loadingPlanPage.ResetFilter();
                DeleteService(serviceName);
                DeleteService(newServiceName);
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_SaveModifications()
        {
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            // Act
            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            loadingPlanPage.ResetFilter();
            var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
            var typeAvant = generalInformationsPage.GetTypeCheck();
            generalInformationsPage.EditLoadingPlanInformationsType();
            var Name = generalInformationsPage.GetLoadingPlanInfo();
            generalInformationsPage.SaveConfirm();
            generalInformationsPage.BackToList();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, Name);
            WebDriver.Navigate().Refresh();
            var typeApres = loadingPlanPage.ClickOnFirstLoadingPlan().GetTypeCheck();
            Assert.AreNotEqual(typeAvant, typeApres, "Les modifications n'ont pas été sauvegardées. Le type reste inchangé après actualisation ");
            var expectedType = generalInformationsPage.GetTypeCheck();
            Assert.AreEqual(typeApres, expectedType, "Le type après modification n'est pas conforme à ce qui est attendu ");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_PictoDeleteService()
        {
            // Prepare
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string route = "ACE - BCN";
            string customerLp = TestContext.Properties["CustomerNamebis"].ToString();
            string loadingPlanName = new Random().Next().ToString();
            string guestName = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string serviceName = "Service1-" + TestContext.Properties["ServiceNameLP"].ToString();
            // Prepare Service
            DateTime fromDate = DateUtils.Now.AddMonths(-1);
            DateTime toDate = DateUtils.Now.AddMonths(+2);

            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            LoadingPlansPage loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            loadingPlanPage.ResetFilter();

            try
            {
                //create a loading Plan
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(toDate, fromDate);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                //add leg and  two services
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                var widthIconDelete = loadingPlanDetailsPage.GetWidthIconDelete();
                var heightIconDelete = loadingPlanDetailsPage.GetHeightIconDelete();
                var widthTdDelete = loadingPlanDetailsPage.GetWidthTDDelete();
                var heightTdDelete = loadingPlanDetailsPage.GetHeightTDDelete();
                bool isIconWithinBounds = loadingPlanDetailsPage.IsIconWithinBounds(widthIconDelete, heightIconDelete, widthTdDelete, heightTdDelete);
                loadingPlanDetailsPage.AddNewService(serviceName);


                Assert.IsTrue(isIconWithinBounds, "la taille de l'icone de suppression du service n'est pas adaptée à l'IHM");
                loadingPlanDetailsPage.BackToList();
            }
            finally
            {
                loadingPlanPage.ResetFilter();
                loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, site, customerLp, toDate);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_StartDate()
        {
            // Prepare
            DateTime startDateInf = DateUtils.Now.AddDays(-12);
            DateTime startDate = DateUtils.Now;
            DateTime startDateSup = DateUtils.Now.AddDays(12);

            // Act
            HomePage homePage = LogInAsAdmin();
            LoadingPlansPage loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                LoadingPlansCreateModalPage loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, CUSTOMER, ROUTE, AIRCRAFT, SITE);
                loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(startDate, startDate);
                LoadingPlansDetailsPage loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                loadingPlanPage = loadingPlanDetailsPage.BackToList();
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.StartDate , startDateInf);
                bool filtred = loadingPlanPage.CheckTotalNumber()==1;
                string loadingPlanFrom = loadingPlanPage.GetFirstLoadingPlanName();
                Assert.AreEqual(loadingPlanName, loadingPlanFrom, "Le nom du loading plan ajouté ne correspond pas");
                Assert.IsTrue(filtred, "les résultats n'accordent pas au filtre appliqué");
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.StartDate, startDateSup);
                filtred = loadingPlanPage.CheckTotalNumber() == 0;
                Assert.IsTrue(filtred, "les résultats n'accordent pas au filtre appliqué");
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.StartDate , startDate);
                filtred = loadingPlanPage.CheckTotalNumber() == 1;
                loadingPlanFrom = loadingPlanPage.GetFirstLoadingPlanName();
                Assert.AreEqual(loadingPlanName, loadingPlanFrom, "Le nom du loading plan ajouté ne correspond pas");
                Assert.IsTrue(filtred, "les résultats n'accordent pas au filtre appliqué");

            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Filter_EndDate()
        {
            // Prepare
            DateTime endDateSup = DateUtils.Now.AddDays(12);
            DateTime endDate = DateUtils.Now;
            DateTime endDateInf = DateUtils.Now.AddDays(-12);

            // Act
            HomePage homePage = LogInAsAdmin();
            LoadingPlansPage loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {
                LoadingPlansCreateModalPage loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, CUSTOMER, ROUTE, AIRCRAFT, SITE);
                loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(endDate, endDate);
                LoadingPlansDetailsPage loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                loadingPlanPage = loadingPlanDetailsPage.BackToList();
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.EndDate, endDateSup);
                bool filtred = loadingPlanPage.CheckTotalNumber() == 1;
                string loadingPlanFrom = loadingPlanPage.GetFirstLoadingPlanName();
                Assert.AreEqual(loadingPlanName, loadingPlanFrom, "Le nom du loading plan ajouté ne correspond pas");
                Assert.IsTrue(filtred, "les résultats n'accordent pas au filtre appliqué");
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.EndDate, endDate);
                filtred = loadingPlanPage.CheckTotalNumber() == 1;
                loadingPlanFrom = loadingPlanPage.GetFirstLoadingPlanName();
                Assert.AreEqual(loadingPlanName, loadingPlanFrom, "Le nom du loading plan ajouté ne correspond pas");
                Assert.IsTrue(filtred, "les résultats n'accordent pas au filtre appliqué");
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.EndDate, endDateInf);
                filtred = loadingPlanPage.CheckTotalNumber() == 0;
                Assert.IsTrue(filtred, "les résultats n'accordent pas au filtre appliqué");

            }
            finally
            {
                DeleteLoadingPlan(loadingPlanName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_GeneralInfo_StartEndDate()
        {
            // Prepare
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string route = TestContext.Properties["RouteLP"].ToString();
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string siteTo = "AGP";
            string loadingPlanName = "LPTEST-" + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNumber = new Random().Next().ToString();
            DateTime lpDateFrom = DateUtils.Now ;
            DateTime lpDateTo = DateUtils.Now.AddMonths(1);

            // Act
            HomePage homePage = LogInAsAdmin();
            //site is active
            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.CheckIfFirstSiteIsActive();

            LoadingPlansPage loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            try
            {

                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customerLp, route, aircraft, site);
                loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(lpDateTo, lpDateFrom);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
             
                // Créer un flight et lui associer le loading plan
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightModalPage = flightPage.FlightCreatePage();
                flightModalPage.FillField_CreatNewFlightWithStartDate(flightNumber, customerLp, aircraft, site, siteTo, null, "00", "23", null, lpDateFrom);
                flightModalPage.CheckLoadingPlan(loadingPlanName);

                //Assert
                bool IsOKExistsLoadingPlan = flightModalPage.IsExistLoadingPlan(loadingPlanName);
                Assert.IsTrue(IsOKExistsLoadingPlan, "LoadingPlan doit être affichée sous le flight pour la même tranche de date utilisée.");

            }
            finally
            {
                //Delete LoadingPlan
                loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
                DeleteLoadingPlan(loadingPlanName);
            }
        }
    }

}