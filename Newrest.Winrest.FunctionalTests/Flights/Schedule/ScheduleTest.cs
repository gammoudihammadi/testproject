using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Schedule;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System;
using System.Security.Policy;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;

namespace Newrest.Winrest.FunctionalTests.Flights
{
    [TestClass]
    public class ScheduleTest : TestBase
    {
        // Impossible de évaluer l’expression, car un frame natif se trouve en haut de la pile des appels.
        private const int _timeout = 60 * 10 * 1000;

        [Priority(0)]
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_SCHE_VerifyServicesForSchedule()
        {
            // Prepare
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();

            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);

            string customer = customerLp.Substring(0, customerLp.IndexOf("("));

            // Arrange
            var homePage = LogInAsAdmin();


            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
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

        // Afficher tous les détails

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_SCHE_Unfold_All()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            var flightNumber = new Random().Next();

            //Arrange
            var homePage = LogInAsAdmin();


            var schedulePage = homePage.GoToFlights_FlightSelectionPage();
            schedulePage.ResetFilter();
            schedulePage.Filter(SchedulePage.FilterType.Site, siteFrom);

            if (schedulePage.CheckTotalNumber() == 0)
            {
                // On créé un vol pour peupler la base des Schedules
                CreateNewFlightWithService(homePage, siteFrom, siteTo, flightNumber);

                schedulePage = homePage.GoToFlights_FlightSelectionPage();
                schedulePage.ResetFilter();
                schedulePage.Filter(SchedulePage.FilterType.Site, siteFrom);
            }

            schedulePage.UnfoldAll();

            // Assert
            Assert.IsFalse(schedulePage.IsFoldAll(), "Les détails des données ne sont pas affichés.");
        }

        // Masquer tous les détails

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_SCHE_Fold_All()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            var flightNumber = new Random().Next();

            // Arrange
            var homePage = LogInAsAdmin();


            var schedulePage = homePage.GoToFlights_FlightSelectionPage();
            schedulePage.ResetFilter();
            schedulePage.Filter(SchedulePage.FilterType.Site, siteFrom);

            if (schedulePage.CheckTotalNumber() == 0)
            {
                // On créé un vol pour peupler la base des Schedules
                CreateNewFlightWithService(homePage, siteFrom, siteTo, flightNumber);

                schedulePage = homePage.GoToFlights_FlightSelectionPage();
                schedulePage.ResetFilter();
                schedulePage.Filter(SchedulePage.FilterType.Site, siteFrom);
            }

            // Act
            schedulePage.UnfoldAll();
            Assert.IsFalse(schedulePage.IsFoldAll(), "Les détails des données ne sont pas affichés.");

            // Assert
            schedulePage.FoldAll();
            Assert.IsTrue(schedulePage.IsFoldAll(), "Les détails des données ne sont pas masqués.");
        }


        // Effectuer des recherches via les filtres - Customer

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_SCHE_Filter_Customers()
        {
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string customerSchedule = TestContext.Properties["CustomerSchedule"].ToString();
            string customerScheduleCode = TestContext.Properties["CustomerScheduleCode"].ToString();
            int flightNumber = new Random().Next();
            string customerScheduleBis = TestContext.Properties["CustomerLP"].ToString();
            HomePage homePage = LogInAsAdmin();
            SchedulePage schedulePage = homePage.GoToFlights_FlightSelectionPage();
            schedulePage.ResetFilter();
            schedulePage.Filter(SchedulePage.FilterType.Site, siteFrom);
            schedulePage.Filter(SchedulePage.FilterType.Customers, customerScheduleCode + " - " + customerSchedule);
            if (schedulePage.CheckTotalNumber() < 10)
            {
                CreateNewFlightWithService(homePage, siteFrom, siteTo, flightNumber);
                schedulePage = homePage.GoToFlights_FlightSelectionPage();
                schedulePage.ResetFilter();
                schedulePage.Filter(SchedulePage.FilterType.Site, siteFrom);
                schedulePage.Filter(SchedulePage.FilterType.Customers, customerScheduleCode + " - " + customerSchedule);
            }
            if (!schedulePage.isPageSizeEqualsTo100())
            {
                schedulePage.PageSize("8");
                schedulePage.PageSize("100");
            }
            bool isVerified = schedulePage.VerifyCustomer(customerScheduleBis);
            Assert.IsTrue(isVerified, string.Format(MessageErreur.FILTRE_ERRONE, "'Customer'"));
        }

        // Effectuer des recherches via les filtres - From date

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_SCHE_Filter_Date()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            var flightNumber = new Random().Next();

            // Arrange
            var homePage = LogInAsAdmin();


            var dateFormat = homePage.GetDateFormatPickerValue();

            var schedulePage = homePage.GoToFlights_FlightSelectionPage();
            schedulePage.ResetFilter();
            schedulePage.Filter(SchedulePage.FilterType.Site, siteFrom);
            schedulePage.Filter(SchedulePage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            schedulePage.Filter(SchedulePage.FilterType.DateTo, DateUtils.Now.AddDays(+15));

            if (schedulePage.CheckTotalNumber() < 10)
            {
                // On créé un vol pour peupler la base des Schedules
                CreateNewFlightWithService(homePage, siteFrom, siteTo, flightNumber);

                schedulePage = homePage.GoToFlights_FlightSelectionPage();
                schedulePage.ResetFilter();
                schedulePage.Filter(SchedulePage.FilterType.Site, siteFrom);
                schedulePage.Filter(SchedulePage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
                schedulePage.Filter(SchedulePage.FilterType.DateTo, DateUtils.Now.AddDays(+15));
            }

            if (!schedulePage.isPageSizeEqualsTo100())
            {
                schedulePage.PageSize("8");
                schedulePage.PageSize("100");
            }

            // Assert
            Assert.IsTrue(schedulePage.IsFromToDateRespected(DateUtils.Now.AddDays(-1).Date, DateUtils.Now.AddDays(+15).Date, dateFormat), String.Format(MessageErreur.FILTRE_ERRONE, "'From/To'"));
        }

        // Effectuer des recherches via les filtres - Site

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_SCHE_Filter_Site()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            var flightNumber = new Random().Next();

            // Arrange
            var homePage = LogInAsAdmin();

            // On créé un vol pour peupler la base des Schedules
            CreateNewFlightWithService(homePage, siteFrom, siteTo, flightNumber);

            // Act
            var schedulePage = homePage.GoToFlights_FlightSelectionPage();
            schedulePage.ResetFilter();
            schedulePage.Filter(SchedulePage.FilterType.Site, siteFrom);

            if (!schedulePage.isPageSizeEqualsTo100())
            {
                schedulePage.PageSize("8");
                schedulePage.PageSize("100");
            }

            // Assert
            Assert.IsTrue(schedulePage.IsFlightNumberPresent(flightNumber.ToString()), String.Format(MessageErreur.FILTRE_ERRONE, "'Site'"));

            // Assert
            schedulePage.Filter(SchedulePage.FilterType.Site, siteTo);
            Assert.IsFalse(schedulePage.IsFlightNumberPresent(flightNumber.ToString()), String.Format(MessageErreur.FILTRE_ERRONE, "'Site'"));
        }

        // Changer le jour de production

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_SCHE_Update_Production_DayD()
        {
            // Prepare
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            var flightNumber = "flight"+new Random().Next();
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteToFlightBob"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);
            string etaHours = "05";
            string etdHours = "23";
            string aircraft = TestContext.Properties["Registration"].ToString();
            int count = 1;
            // Arrange
            var homePage = LogInAsAdmin();
            var servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                // Act
                // créé un service
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, serviceCategory);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                // add a price
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, dateFrom, dateTo);
                servicePage = pricePage.BackToList();
                // créé flight 
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                // associate a service for the flight
                var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
                editPage.AddGuestType();
                editPage.AddService(serviceName);
                editPage.CloseViewDetails();
                // go to schelude
                var schedulePage = homePage.GoToFlights_FlightSelectionPage();
                schedulePage.ResetFilter();
                schedulePage.Filter(SchedulePage.FilterType.Site, site);
                schedulePage.Filter(SchedulePage.FilterType.DateFrom, dateFrom);
                schedulePage.UnfoldAll();
                if (schedulePage.isServiceAtDayD()) 
                {
                    schedulePage.SetProdDayD_1();
                    var initValue = schedulePage.GetNomnbreServiceAtDayD();
                    schedulePage.SetProdDayD();
                    var newValue = schedulePage.GetNomnbreServiceAtDayD();
                    //Asset
                    Assert.AreNotEqual(initValue, newValue, "La valeur du Production Day n'a pas été modifiée.");
                    Assert.AreEqual(count, newValue, "La valeur du Production Day n'a pas été modifiée.");
                }
                else 
                {
                    var initValue = schedulePage.GetNomnbreServiceAtDayD();
                    schedulePage.SetProdDayD();
                    var newValue = schedulePage.GetNomnbreServiceAtDayD();
                    // Assert
                    Assert.AreNotEqual(initValue, newValue, "La valeur du Production Day n'a pas été modifiée.");
                    Assert.AreEqual(count, newValue, "La valeur du Production Day n'a pas été modifiée.");
                }
            }
            finally
            {
                // delete a service 
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                //Delete Flight
                var flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_SCHE_Update_Production_DayD_1()
        {
            // Prepare
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            var flightNumber = "flight" + new Random().Next();
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteToFlightBob"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);
            string etaHours = "05";
            string etdHours = "23";
            string aircraft = TestContext.Properties["Registration"].ToString();
            int count = 1;
            // Arrange
            var homePage = LogInAsAdmin();
            var servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                // Act
                // créé un service
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, serviceCategory);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                // add a price
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, dateFrom, dateTo);
                servicePage = pricePage.BackToList();
                // créé flight 
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo, null, etaHours, etdHours);
                // associate a service for the flight
                var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
                editPage.AddGuestType();
                editPage.AddService(serviceName);
                editPage.CloseViewDetails();
                // go to schelude
                var schedulePage = homePage.GoToFlights_FlightSelectionPage();
                schedulePage.ResetFilter();
                schedulePage.Filter(SchedulePage.FilterType.Site, site);
                schedulePage.Filter(SchedulePage.FilterType.DateFrom, dateFrom);
                schedulePage.UnfoldAll();
                if (schedulePage.isServiceAtDayD())
                {
                    var initValue = schedulePage.GetNomnbreServiceAtDayD_1();
                    schedulePage.SetProdDayD_1();
                    var newValue = schedulePage.GetNomnbreServiceAtDayD_1();
                    //Asset
                    Assert.AreNotEqual(initValue, newValue, "La valeur du Production DayD_1 n'a pas été modifiée.");
                    Assert.AreEqual(count, newValue, "La valeur du Production DayD_1 n'a pas été modifiée.");
                }
                else
                {
                    schedulePage.SetProdDayD();
                    var initValue = schedulePage.GetNomnbreServiceAtDayD_1();
                    schedulePage.SetProdDayD_1();
                    var newValue = schedulePage.GetNomnbreServiceAtDayD_1();
                    // Assert
                    Assert.AreNotEqual(initValue, newValue, "La valeur du Production DayD_1 n'a pas été modifiée.");
                    Assert.AreEqual(count, newValue, "La valeur du Production DayD_1 n'a pas été modifiée.");
                }
            }
            finally
            {
                // delete a service 
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                //Delete Flight
                var flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
            }
        }
        // Cocher/décocher la production

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_SCHE_Update_Production()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            var flightNumber = new Random().Next();

            // Arrange
            var homePage = LogInAsAdmin();


            // On créé un vol pour peupler la base des Schedules
            CreateNewFlightWithService(homePage, siteFrom, siteTo, flightNumber);

            // Act
            var schedulePage = homePage.GoToFlights_FlightSelectionPage();
            schedulePage.ResetFilter();
            schedulePage.Filter(SchedulePage.FilterType.Site, siteFrom);

            if (!schedulePage.isPageSizeEqualsTo100())
            {
                //schedulePage.PageSize("8");
                schedulePage.PageSize("100");
            }

            schedulePage.UnfoldAll();

            schedulePage.ChangeIsProduced(false);
            var notProducedState = schedulePage.GetServiceNonProducedValue();
            schedulePage.ChangeIsProduced(true);
            var producedState = schedulePage.GetServiceNonProducedValue();
            schedulePage.ChangeIsProduced(false);
            var backNotProducedState = schedulePage.GetServiceNonProducedValue();

            //Assert
            Assert.AreNotEqual(notProducedState, producedState, "La valeur de la Production n'a pas été modifiée.");
            Assert.AreEqual(notProducedState, backNotProducedState, "La valeur de la Production n'a pas été modifiée.");

        }

        public void CreateNewFlightWithService(HomePage homePage, string siteFrom, string siteTo, int flightNumber)
        {
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var customer = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber.ToString(), customer, aircraft, siteFrom, siteTo, null, "00", "01");

            //Edit the first flight
            var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.AddGuestType();
            editPage.AddService(serviceName);
            editPage.CloseViewDetails();
        }
    }
}