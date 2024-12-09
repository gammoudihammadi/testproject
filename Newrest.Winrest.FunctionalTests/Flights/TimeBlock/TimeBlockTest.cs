using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.TimeBlock;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Tablet;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using static Newrest.Winrest.FunctionalTests.PageObjects.TabletApp.TimeBlockTabletAppPage;

namespace Newrest.Winrest.FunctionalTests.Flights
{
    [TestClass]
    public class TimeBlockTest : TestBase
    {
        // Impossible de évaluer l’expression, car un frame natif se trouve en haut de la pile des appels.
        private const int _timeout = 60 * 10 * 1000;

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Filter_Site()
        {
            //Prepare
            string airportFrom = TestContext.Properties["Site"].ToString();
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNo = new Random().Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            //Créer un Flight
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            FlightCreateModalPage create = flightPage.FlightCreatePage();
            create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo);
            // Act
            //Etre sur l'index des Time Block et avoir des données
            TimeBlockPage timeBlockPage = homePage.GoToFlight_TimeBlockPage();
            timeBlockPage.ResetFilters();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
            if (!timeBlockPage.isPageSizeEqualsTo100())
            {
                timeBlockPage.PageSize("100");
            }
            try
            {
                //Parfois MAD est siteFrom, et parfois siteTo
                ReadOnlyCollection<IWebElement> elementsSiteFrom = WebDriver.FindElements(By.XPath("//*[@id='list-item-with-action']/table/tbody/tr[*]/td[5]"));
                ReadOnlyCollection<IWebElement> elementsSiteTo = WebDriver.FindElements(By.XPath("//*[@id='list-item-with-action']/table/tbody/tr[*]/td[6]"));
                Assert.AreEqual(elementsSiteFrom.Count, elementsSiteTo.Count);
                for (int i = 0; i<elementsSiteFrom.Count;i++)
                {
                    if (elementsSiteFrom[i].Text == airportFrom)
                    {
                        continue;
                    }
                    if (elementsSiteTo[i].Text == airportFrom)
                    {
                        continue;
                    }
                    Assert.Fail("le filtre sur le site ne s'applique pas : MAD non dans from " + elementsSiteFrom[i].Text + " et non dans to " + elementsSiteTo[i].Text);
                }
            }
            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNo);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }
        
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Filter_Search()
        {
            //Prepare
            string airportFrom = TestContext.Properties["Site"].ToString();
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNo = new Random().Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            //Créer un Flight
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            FlightCreateModalPage create = flightPage.FlightCreatePage();
            create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo);
            // Act
            //Etre sur l'index des Time Block et avoir des données
            try
            {
                TimeBlockPage timeBlockPage = homePage.GoToFlight_TimeBlockPage();
                timeBlockPage.ResetFilters();
                //Effectuer un filtre Search
                timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNo);
                timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
                bool verifyName = timeBlockPage.GetResultFilght().All(names => names.Contains(flightNo));
                Assert.IsTrue(verifyName, "le filtre search  ne s'applique pas");
               
            }

            finally
            {
                flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteModalPage = flightPage.ClickMassiveDelete();
                massiveDeleteModalPage.SetFlightName(flightNo);
                massiveDeleteModalPage.ClickSearchButton();
                massiveDeleteModalPage.ClickSelectAllButton();
                massiveDeleteModalPage.Delete();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Filter_MajorFlightsOnly()
        {
            //Prepare
            string airportFrom = TestContext.Properties["Site"].ToString();
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNo = new Random().Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            //Créer un Flight
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            FlightCreateModalPage create = flightPage.FlightCreatePage();
            create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo);
            // Act
            //Etre sur l'index des Time Block et avoir des données
            TimeBlockPage timeBlockPage = homePage.GoToFlight_TimeBlockPage();
            timeBlockPage.ResetFilters();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNo);
            timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
            timeBlockPage.Filter(TimeBlockPage.FilterType.MajorFlightsOnly, false);
            Assert.AreEqual(1, timeBlockPage.CheckTotalNumber());

            timeBlockPage.Filter(TimeBlockPage.FilterType.MajorFlightsOnly, true);
            Assert.AreEqual(0, timeBlockPage.CheckTotalNumber());
            timeBlockPage.Filter(TimeBlockPage.FilterType.MajorFlightsOnly, false);
            //Appliquer le CheckBox Major Flights Only
            timeBlockPage.SetFirstMajorFlightsOnly(true);
            timeBlockPage.Filter(TimeBlockPage.FilterType.MajorFlightsOnly, true);
            Assert.AreEqual(1, timeBlockPage.CheckTotalNumber());
            Assert.AreEqual(flightNo, timeBlockPage.GetFirstFlightNumber());
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Filter_HideCancelledFlights()
        {
            //Prepare
            string airportFrom = TestContext.Properties["Site"].ToString();
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNo = new Random().Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            //Act
            //Créer un Flight
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            FlightCreateModalPage create = flightPage.FlightCreatePage();
            create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNo);
            flightPage.SetNewState("C");
            //Etre sur l'index des Time Block et avoir des données
            TimeBlockPage timeBlockPage = homePage.GoToFlight_TimeBlockPage();
            timeBlockPage.ResetFilters();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
            timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNo);
            //Appliquer le CheckBox Hide Cancelled Flights
            timeBlockPage.Filter(TimeBlockPage.FilterType.HideCancelledFlights, false);
            //Les données sont filtrées en fonction du filtre appliquer
            Assert.AreEqual(1, timeBlockPage.CheckTotalNumber(), "HideCancelledFlights false ne remonte pas un Cancelled Flight");
            //Appliquer le CheckBox Hide Cancelled Flights
            timeBlockPage.Filter(TimeBlockPage.FilterType.HideCancelledFlights, true);
            //Les données sont filtrées en fonction du filtre appliquer
            Assert.AreEqual(0, timeBlockPage.CheckTotalNumber(), "HideCancelledFlights true remonte un Cancelled Flight");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Filter_ETDFrom()
        {
            //Prepare
            string airportFrom = TestContext.Properties["Site"].ToString();
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flight1No = "Flight1" + new Random().Next().ToString();
            string flight2No = "Flight2" + new Random().Next().ToString();
            string flight3No = "Flight3" + new Random().Next().ToString();

            string etdFrom = "0900AM";
            string etdTo = "1100AM";
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Créer un Flight
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            try
            {
                FlightCreateModalPage create = flightPage.FlightCreatePage();
                create.FillField_CreatNewFlight(flight1No, customer, aircraft, airportFrom, airportTo, null, "06", "08");

                //Etre sur l'index des Time Block et avoir des données
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1No);
                TimeBlockPage timeBlockPage = homePage.GoToFlight_TimeBlockPage();
                timeBlockPage.ResetFilters();
                timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flight1No);

                //Appliquer filtre ETD From 
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETA, etdFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETD, etdTo);
                //Les données sont filtrées en fonction du filtre appliquer
                var totalNumber = timeBlockPage.CheckTotalNumber();
                Assert.AreEqual(0, totalNumber, " La flight 1 apparait aprés le filtre.");

                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
                create = flightPage.FlightCreatePage();
                create.FillField_CreatNewFlight(flight2No, customer, aircraft, airportFrom, airportTo, null, "06", "09");
                //Etre sur l'index des Time Block et avoir des données
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1No);
                timeBlockPage = homePage.GoToFlight_TimeBlockPage();
                timeBlockPage.ResetFilters();
                timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flight2No);

                //Appliquer filtre ETD From 
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETA, etdFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETD, etdTo);
                //Les données sont filtrées en fonction du filtre appliquer
                Assert.AreEqual(1, timeBlockPage.CheckTotalNumber(), " La flight 2 n'apparait pas.");

                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
                create = flightPage.FlightCreatePage();
                create.FillField_CreatNewFlight(flight3No, customer, aircraft, airportFrom, airportTo, null, "06", "10");

                //Etre sur l'index des Time Block et avoir des données
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight3No);
                timeBlockPage = homePage.GoToFlight_TimeBlockPage();
                timeBlockPage.ResetFilters();
                timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flight3No);

                //Appliquer filtre ETD From 
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETA, etdFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETD, etdTo);
                //Les données sont filtrées en fonction du filtre appliquer
                Assert.AreEqual(1, timeBlockPage.CheckTotalNumber(), " La flight 3 n'apparait pas.");
            }
            finally
            {
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1No);
                flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
                flightPage.MassiveDeleteMenus(flight1No, airportFrom, null);

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight2No);
                flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
                flightPage.MassiveDeleteMenus(flight2No, airportFrom, null);

                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight3No);
                flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
                flightPage.MassiveDeleteMenus(flight3No, airportFrom, null);
            }

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Filter_ETDTo()
        {
            //Prepare
            string airportFrom = TestContext.Properties["Site"].ToString();
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flight1No = "Flight1" + new Random().Next().ToString();
            string flight2No = "Flight2" + new Random().Next().ToString();
            string flight3No = "Flight3" + new Random().Next().ToString();

            string etdFrom = "0000AM";
            string etdTo = "1000AM";
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Créer un Flight
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            try
            {
                FlightCreateModalPage create = flightPage.FlightCreatePage();
                create.FillField_CreatNewFlight(flight1No, customer, aircraft, airportFrom, airportTo, null, "00", "09");
                //flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNo);

                //Etre sur l'index des Time Block et avoir des données
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1No);
                TimeBlockPage timeBlockPage = homePage.GoToFlight_TimeBlockPage();
                timeBlockPage.ResetFilters();
                timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flight1No);

                //Appliquer filtre ETD From et To 00:00-12:00
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETA, etdFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETD, etdTo);
                //Les données sont filtrées en fonction du filtre appliquer
                var totalNumber = timeBlockPage.CheckTotalNumber();
                Assert.AreEqual(1, totalNumber, " La flight 1 n'apparait pas.");

                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
                create = flightPage.FlightCreatePage();
                create.FillField_CreatNewFlight(flight2No, customer, aircraft, airportFrom, airportTo, null, "00", "10");
                //Etre sur l'index des Time Block et avoir des données
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1No);
                timeBlockPage = homePage.GoToFlight_TimeBlockPage();
                timeBlockPage.ResetFilters();
                timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flight2No);

                //Appliquer filtre ETD From et To 00:00-12:00
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETA, etdFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETD, etdTo);
                //Les données sont filtrées en fonction du filtre appliquer
                Assert.AreEqual(1, timeBlockPage.CheckTotalNumber(), " La flight 2 n'apparait pas.");

                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
                create = flightPage.FlightCreatePage();
                create.FillField_CreatNewFlight(flight3No, customer, aircraft, airportFrom, airportTo, null, "00", "11");

                //Etre sur l'index des Time Block et avoir des données
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight3No);
                timeBlockPage = homePage.GoToFlight_TimeBlockPage();
                timeBlockPage.ResetFilters();
                timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flight3No);

                //Appliquer filtre ETD From et To 00:00-12:00
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETA, etdFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.HourETD, etdTo);
                //Les données sont filtrées en fonction du filtre appliquer
                Assert.AreEqual(0, timeBlockPage.CheckTotalNumber(), " La flight 3 apparait.");

            }
            finally
            {
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1No);
                flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
                flightPage.MassiveDeleteMenus(flight1No, airportFrom, null);

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight2No);
                flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
                flightPage.MassiveDeleteMenus(flight2No, airportFrom, null);

                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight3No);
                flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
                flightPage.MassiveDeleteMenus(flight3No, airportFrom, null);
            }

        }
        
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Filter_Customer()
        {
            //Prepare
            string airportFrom = TestContext.Properties["Site"].ToString();
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNo = new Random().Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            //Act
            //Créer un Flight
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            FlightCreateModalPage create = flightPage.FlightCreatePage();
            create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNo);
            //Etre sur l'index des Time Block et avoir des données
            TimeBlockPage timeBlockPage = homePage.GoToFlight_TimeBlockPage();
            timeBlockPage.ResetFilters();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
            timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNo);

            //Appliquer le filtre Customer
            timeBlockPage.Filter(TimeBlockPage.FilterType.Customer, customer);
            Assert.AreEqual(1, timeBlockPage.CheckTotalNumber());
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_ResetFilter()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            homePage.Navigate();
            TimeBlockPage timeBlockPage = homePage.GoToFlight_TimeBlockPage();
            timeBlockPage.ResetFilters();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Site, "ACE");
            int avant = timeBlockPage.CheckTotalNumber();

            timeBlockPage.Filter(TimeBlockPage.FilterType.Site, site);
            //search,
            timeBlockPage.Filter(TimeBlockPage.FilterType.Search, "TEST");
            //customer,
            timeBlockPage.Filter(TimeBlockPage.FilterType.Customer, customer);
            // dates - from et to-,
            timeBlockPage.Filter(TimeBlockPage.FilterType.DateFrom, DateUtils.Now.AddDays(-15));
            timeBlockPage.Filter(TimeBlockPage.FilterType.DateTo, DateUtils.Now.AddDays(15));
            //ETD.
            timeBlockPage.Filter(TimeBlockPage.FilterType.HourETA, "1205AM");
            timeBlockPage.Filter(TimeBlockPage.FilterType.HourETD, "1205PM");
            // major, 
            timeBlockPage.Filter(TimeBlockPage.FilterType.MajorFlightsOnly, true);
            timeBlockPage.Filter(TimeBlockPage.FilterType.HideCancelledFlights, true);

            int apres = timeBlockPage.CheckTotalNumber();
            // apres à 0 car Search à "TEST"
            Assert.AreNotEqual(avant, apres, "Aucune donnée après un ResetFilter");

            //Cliquer sur Reset Filtrer
            timeBlockPage.ResetFilters();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Site, "ACE");
            int maintenant = timeBlockPage.CheckTotalNumber();
            Assert.AreEqual(avant, maintenant, "ResetFilter KO");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Tablet_Filter_Search()
        {
            var homePage = LogInAsAdmin();
            var timeBlockpage = homePage.GoToFlight_TimeBlockPage();
            timeBlockpage.Filter(TimeBlockPage.FilterType.Site, "ACE");
            var firstFlightNumber = timeBlockpage.GetFirstFlightNumber();
            timeBlockpage.Filter(TimeBlockPage.FilterType.Search, firstFlightNumber);
            Assert.IsTrue(timeBlockpage.VerifyFilterByFlightNumbre(firstFlightNumber), "Erreur de filtrage par flight number");
        }
        [TestMethod]
        [Priority(2)]
        [Timeout(_timeout)]
        public void FL_TB_SetTimeBlockFlightTypeParameter()
        {
            // Prepare
            string couleur = "Gray";
            string flightType = "Regular";
            var site = "ACE";
            //login as admin
            var homePage = LogInAsAdmin();

            var tabletPage = homePage.GoToParametres_Tablet();
            tabletPage.ClickFlightTypeTab();
            //check if flight type international exist
            var exist = tabletPage.VerifyInternationalExistOnSite(flightType, site);
            // if no create the flight type international
            if (!exist)
            {
                tabletPage.CreateNewFlightType(flightType, couleur, site);
            }
            // else set the color gray
            tabletPage.EditFlightTypeColor(flightType, couleur, site);
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            var flightNo = flightPage.GetFirstFlightNumber();
            var flightDetailsPage = flightPage.EditFirstFlight(flightNo);
            // combobox Flight Type non grisé : P
            flightDetailsPage.UnsetState("P");
            flightDetailsPage.SetNewState("P");

            flightDetailsPage.SetTabletFlightType(flightType);
            //verify flight type of the flight
            homePage.Navigate();
            var tabletAppPage = homePage.GotoTabletApp();
            var timeBlockTabletAppPage = tabletAppPage.GotoTabletApp_TimeBlock();
            timeBlockTabletAppPage.FilterByFlightNumber(flightNo);

            Assert.IsTrue(timeBlockTabletAppPage.VerifyFlightType("gray", flightType, true), "erreur d'affectation de flight type");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Tablet_Filter_FlightType()
        {
            // Prepare
            var flightNo = new Random().Next().ToString();
            string couleur = "Gray";
            string flightType = "Regular";// pas de choix International pour ce nouveau vol ?!?
            var siteFrom = "ACE";
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customer = "TVS";// TestContext.Properties["CustomerSchedule"].ToString();
            //login as admin
            var homePage = LogInAsAdmin();
            var tabletPage = homePage.GoToParametres_Tablet();
            tabletPage.ClickFlightTypeTab();
            //check if flight type international exist
            var exist = tabletPage.VerifyInternationalExistOnSite(flightType, siteTo);
            // if no create the flight type international
            if (!exist)
            {
                tabletPage.CreateNewFlightType(flightType, couleur, siteTo);
            }
            // else set the color gray
            tabletPage.EditFlightTypeColor(flightType, couleur, siteTo);
            //create new flight with flight type international
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var createFlightpage = flightPage.FlightCreatePage();
            createFlightpage.FillField_CreatNewFlightWithFlightType(flightNo, customer, aircraft, siteFrom, siteTo, flightType);
            FlightDetailsPage filghetDetail = flightPage.EditFirstFlight(flightNo);
            filghetDetail.SetTabletFlightType(flightType);
            filghetDetail.CloseModal();

            homePage.Navigate();
            //verify 
            var tabletAppPage = homePage.GotoTabletApp();
            var timeBlockTabletAppPage = tabletAppPage.GotoTabletApp_TimeBlock();
            timeBlockTabletAppPage.TimeBlockResetFilters();
            timeBlockTabletAppPage.FilterByFlightType(flightType);
            Assert.IsTrue(timeBlockTabletAppPage.VerifyFlightType("gray", flightType), "le flight type n'est pas appliqué sur le vol");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_ScheduleWorshop()
        {
            //Prepare
            string airportFrom = "ACE";
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = "$$ - CAT Genérico";
            string flightNo = new Random().Next().ToString();
            var j2 = DateTime.Now.AddDays(2);
            var j1 = DateTime.Now.AddDays(1);

            var homePage = LogInAsAdmin();


            ParametersTablet tabletPage = homePage.GoToParametres_Tablet();
            tabletPage.WorkshopsTab();
            tabletPage.AllWorkshopIsTimeBlock("History");
            tabletPage.WorkshopExistOnSite(airportFrom, "History");
            tabletPage.ShowOnTabletByLabel("History");


            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            FlightCreateModalPage create = flightPage.FlightCreatePage();
            create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo, null, "00", "06", null, j2);

            var timeBlockPage = homePage.GoToFlight_TimeBlockPage();
            timeBlockPage.ResetFilters();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
            timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNo);
            timeBlockPage.Filter(TimeBlockPage.FilterType.DateTo, j2);

            timeBlockPage.Schedule();
            timeBlockPage.SetFirstWorkshopParams(j1.ToString("dd/MM/yyyy"), "04:00", j1.ToString("dd/MM/yyyy"), "02:00");
            timeBlockPage.TimeBeforeEtdKtd();
            Assert.IsTrue(timeBlockPage.VerifyWorkshopDaysAndTime());
            homePage.Navigate();
            var tabletAppPage = homePage.GotoTabletApp();
            var timeBlockTabletApp = tabletAppPage.GotoTabletApp_TimeBlock();
            timeBlockTabletApp.Filter(FilterType.Search, flightNo);
            timeBlockTabletApp.SetDate(j2);
            Assert.IsTrue(timeBlockPage.VerifyNotLate(), "la couleur des horaires des workshop est rouge");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_ScheduleWorshopDelay()
        {
            //Prepare
            string airportFrom = "ACE";
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = "$$ - CAT Genérico";
            string flightNo = new Random().Next().ToString();
            var j = DateTime.Now;
            var jMinus1 = DateTime.Now.AddDays(-1);

            var homePage = LogInAsAdmin();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            FlightCreateModalPage create = flightPage.FlightCreatePage();
            create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo, null, "00", "06", null, j);
            var timeBlockPage = homePage.GoToFlight_TimeBlockPage();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
            timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNo);
            timeBlockPage.Filter(TimeBlockPage.FilterType.DateTo, j);

            timeBlockPage.Schedule();
            timeBlockPage.SetFirstWorkshopParams(jMinus1.ToString("dd/MM/yyyy"), "04:00", jMinus1.ToString("dd/MM/yyyy"), "02:00");
            timeBlockPage.TimeBeforeEtdKtd();
            Assert.IsTrue(timeBlockPage.VerifyWorkshopDaysAndTime());
            homePage.Navigate();
            var tabletAppPage = homePage.GotoTabletApp();
            var timeBlockTabletApp = tabletAppPage.GotoTabletApp_TimeBlock();
            timeBlockTabletApp.Filter(FilterType.Search, flightNo);
            timeBlockTabletApp.SetDate(j);
            Assert.IsTrue(!timeBlockPage.VerifyNotLate(), "la couleur des horaires des workshop est noir");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_SetTimeBlockWorkshopsParameters()
        {
            //prepare 
            string ensamblaje = "Ensamblaje";
            string corte = "Corte";
            string ensambl = "Ensambl";
            var homePage = LogInAsAdmin();

            var applicationSettings = homePage.GoToApplicationSettings();
            var parametersProduction = applicationSettings.GoToParameters_ProductionPage();
            parametersProduction.WorkshopTab();
            //verify ensamblaje et corte existent
            var exist = parametersProduction.VerifyWorkshopExist(ensamblaje, corte);
            if (!exist)
            {
                parametersProduction.CreateNewWorkshop(ensamblaje);
                parametersProduction.CreateNewWorkshop(corte);
            }
            var paramTablet = applicationSettings.GoToParametres_Tablet();
            paramTablet.WorkshopsTab();
            paramTablet.ShowOnTabletByLabel(ensambl, corte);
            paramTablet.AffecterOrdre(ensambl, 1);
            paramTablet.AffecterOrdre(corte, 2);
            var timeBlockPage = homePage.GoToFlight_TimeBlockPage();
            var work1 = timeBlockPage.VerifyColumnsExist(ensamblaje);
            var work2 = timeBlockPage.VerifyColumnsExist(corte);
            Assert.IsTrue(work1, "workshop ensamblaje n'existe pas dans flights time block");
            Assert.IsTrue(work2, "workshop corte n'existe pas dans flights time block");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_StartAndEndWorkshopTimes()
        {
            //FIXME
            // dépend de FL_TB_ScheduleWorshop
            //Prepare
            string airportFrom = "ACE";
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = "$$ - CAT Genérico";
            string flightNo = new Random().Next().ToString();
            var j2 = DateTime.Now.AddDays(2);
            var j1 = DateTime.Now.AddDays(1);
            var j0 = DateTime.Now;
            var jm1 = DateTime.Now.AddDays(-1);
            //login as admin
            var homePage = LogInAsAdmin();


            //go to flights page
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            //create new flight
            FlightCreateModalPage create = flightPage.FlightCreatePage();
            // date à demain, pour avoir de la donnée aujourd'hui (par défaut Schedule à J-1)
            create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo, null, "00", "06", null, j1);

            //go to timeblock page
            var timeBlockPage = homePage.GoToFlight_TimeBlockPage();
            timeBlockPage.ResetFilters();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
            timeBlockPage.Filter(TimeBlockPage.FilterType.Customer, customer);
            timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNo);
            // par défaut en Schedule en J-1
            timeBlockPage.Filter(TimeBlockPage.FilterType.DateFrom, j1);
            timeBlockPage.Filter(TimeBlockPage.FilterType.DateTo, j1);

            //schedule workshops
            timeBlockPage.Schedule();
            timeBlockPage.SetFirstWorkshopParams(j1.ToString("dd/MM/yyyy"), "04:00", j1.ToString("dd/MM/yyyy"), "02:00");
            timeBlockPage.SetSecondWorkshopParams(j1.ToString("dd/MM/yyyy"), "04.30", j1.ToString("dd/MM/yyyy"), "02.30");
            timeBlockPage.TimeBeforeEtdKtd();

            // go to tablet app
            homePage.Navigate();
            var tabletApp = homePage.GotoTabletApp();
            tabletApp.SetSite(airportFrom);
            var timeBlockTabletApp = tabletApp.GotoTabletApp_TimeBlock();
            timeBlockTabletApp.Filter(FilterType.Search, flightNo);
            timeBlockTabletApp.SetDate(j0);
            WebDriver.Navigate().Refresh();
            //start workshops
            var firstWorkshopStartColor = timeBlockTabletApp.StartFirstWorkshop();
            Assert.AreEqual("Green", firstWorkshopStartColor, "Couleur de bouton n'est pas vert");
            Assert.IsTrue(timeBlockTabletApp.VerifyState(State.STARTED), "l'etat n'est pas passé à Started");

            var secondWorkshopEndColor = timeBlockTabletApp.EndFirstWorkshop();
            Assert.AreEqual("Green", secondWorkshopEndColor, "Couleur de bouton n'est pas vert");

            timeBlockTabletApp.StartSecondWorkshop();
            timeBlockTabletApp.EndSecondWorkshop();
            Assert.IsTrue(timeBlockTabletApp.VerifyState(State.DONE), "l'etat n'est pas passé à Done");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_StartAndEndWorkshopTimesWithDelay()
        {
            //Prepare
            string airportFrom = "ACE";
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = "$$ - CAT Genérico";
            string flightNo = new Random().Next().ToString();
            var j3 = DateTime.Now.AddDays(3);
            var j2 = DateTime.Now.AddDays(2);
            var j1 = DateTime.Now.AddDays(1);
            var j0 = DateTime.Now;
            var jm1 = DateTime.Now.AddDays(-1);
            var jm2 = DateTime.Now.AddDays(-2);
            var jm3 = DateTime.Now.AddDays(-3);
            //login as admin
            var homePage = LogInAsAdmin();

            //go to flights page
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            //create new flight
            FlightCreateModalPage create = flightPage.FlightCreatePage();
            // date à demain, pour avoir de la donnée aujourd'hui (par défaut Schedule à J-1)
            create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo, null, "00", "06", null, j1);

            //go to timeblock page
            var timeBlockPage = homePage.GoToFlight_TimeBlockPage();
            timeBlockPage.ResetFilters();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
            timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNo);
            timeBlockPage.Filter(TimeBlockPage.FilterType.DateFrom, j1);
            timeBlockPage.Filter(TimeBlockPage.FilterType.DateTo, j1);

            //schedule workshops
            timeBlockPage.Schedule();
            //Attention il faut les 2 workshops Ensemblaje et Corte
            // plage de date entre 1 et 3
            timeBlockPage.SetFirstWorkshopParams(j1.ToString("dd/MM/yyyy"), "04:00", j3.ToString("dd/MM/yyyy"), "02:00");
            timeBlockPage.SetSecondWorkshopParams(j1.ToString("dd/MM/yyyy"), "04.30", j3.ToString("dd/MM/yyyy"), "02.30");
            timeBlockPage.TimeBeforeEtdKtd();

            // go to tablet app
            homePage.Navigate();
            var tabletApp = homePage.GotoTabletApp();
            tabletApp.SetSite(airportFrom);
            var timeBlockTabletApp = tabletApp.GotoTabletApp_TimeBlock();
            timeBlockTabletApp.Filter(FilterType.Search, flightNo);

            timeBlockTabletApp.SetDate(j0);
            //WebDriver.Navigate().Refresh();
            //start workshops
            var firstWorkshopStartColor = timeBlockTabletApp.StartFirstWorkshop();
            Assert.AreEqual("Green", firstWorkshopStartColor, "Couleur de bouton n'est pas rouge");
            Assert.IsTrue(timeBlockTabletApp.VerifyState(State.STARTED), "l'etat n'est pas passé à Started");
            var secondWorkshopEndColor = timeBlockTabletApp.EndFirstWorkshop();
            Assert.AreEqual("Green", secondWorkshopEndColor, "Couleur de bouton n'est pas rouge");
            timeBlockTabletApp.StartSecondWorkshop();
            timeBlockTabletApp.EndSecondWorkshop();
            Assert.IsTrue(timeBlockTabletApp.VerifyState(State.DONE), "l'etat n'est pas passé à Done");

            timeBlockTabletApp.SetDate(jm1);
            //WebDriver.Navigate().Refresh();
            //start workshops
            firstWorkshopStartColor = timeBlockTabletApp.FirstWorkshopStarted();
            Assert.AreEqual("Green", firstWorkshopStartColor, "Couleur de bouton n'est pas rouge (cas 2)");
            Assert.IsTrue(timeBlockTabletApp.VerifyState(State.DONE), "l'etat n'est pas passé à DONE (cas 2)");
            secondWorkshopEndColor = timeBlockTabletApp.FirstWorkshopEnded();
            Assert.AreEqual("Green", secondWorkshopEndColor, "Couleur de bouton n'est pas rouge (cas 2)");
            timeBlockTabletApp.SecondWorkshopStarted();
            timeBlockTabletApp.SecondWorkshopEnded();
            Assert.IsTrue(timeBlockTabletApp.VerifyState(State.DONE), "l'etat n'est pas passé à Done (cas 2)");

            timeBlockTabletApp.SetDate(jm2);
            //WebDriver.Navigate().Refresh();
            //start workshops
            firstWorkshopStartColor = timeBlockTabletApp.FirstWorkshopStarted();
            Assert.AreEqual("Green", firstWorkshopStartColor, "Couleur de bouton n'est pas rouge (cas 3)");
            Assert.IsTrue(timeBlockTabletApp.VerifyState(State.DONE), "l'etat n'est pas passé à DONE (cas 3)");
            secondWorkshopEndColor = timeBlockTabletApp.FirstWorkshopEnded();
            Assert.AreEqual("Green", secondWorkshopEndColor, "Couleur de bouton n'est pas rouge (cas 3)");
            timeBlockTabletApp.SecondWorkshopStarted();
            timeBlockTabletApp.SecondWorkshopEnded();
            Assert.IsTrue(timeBlockTabletApp.VerifyState(State.DONE), "l'etat n'est pas passé à Done");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Tablet_Filter_Customer()
        {
            var customer = "TVS";

            var homePage = LogInAsAdmin();

            var tabletApp = homePage.GotoTabletApp();
            var tabletAppTimeBlock = tabletApp.GotoTabletApp_TimeBlock();
            tabletAppTimeBlock.TimeBlockResetFilters();

            tabletAppTimeBlock.Filter(FilterType.Customers, customer);
            Assert.IsTrue(tabletAppTimeBlock.VerifyCustomers(customer),"le filtre ne s'applique pas");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_BulkChange()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();

            // Arrange

            var homePage = LogInAsAdmin();


            //Act
            TimeBlockPage timeBlock = homePage.GoToFlight_TimeBlockPage();
            timeBlock.ResetFilters();

            //Avoir un site avec un customer et des vols
            timeBlock.Filter(TimeBlockPage.FilterType.Site, site);
            timeBlock.Filter(TimeBlockPage.FilterType.Customer, customer);
            Assert.AreNotEqual(0, timeBlock.CheckTotalNumber());

            //1. Survoler les ... et cliquer sur Bulk Change
            BulkChangeModal BCmodal = timeBlock.BulkChange();

            //2. Une pop'up s'ouvre
            //et qualifier quelques champs (jours et horaires)
            BCmodal.Fill("5", "1236AM", "2", "0335PM", "2", "0238AM", "1", "0735PM", true);
            //et cliquer Validate
            BCmodal.Submit();

            //3. Vérifier que les modifs soient effectif
            //(changement de couleur) sur l'index des Time Block
            // Les modifications sont prises en compte sur l'index
            BCmodal.CheckFirstFlight("5", "00:36:00", "2", "15:35:00", "2", "02:38:00", "1", "19:35:00", true);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_ChargementLorsModif_Jour_Heure()
        {
            //Prepare
            string airportFrom = TestContext.Properties["Site"].ToString();
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNo = new Random().Next().ToString();
            var j2 = DateTime.Now.AddDays(2);
            var j1 = DateTime.Now.AddDays(1);
            // Arrange
            var homePage = LogInAsAdmin();


            //Créer un Flight
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);

            try
            {
                FlightCreateModalPage create = flightPage.FlightCreatePage();
                create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo);
                // Act
                //Etre sur l'index des Time Block et avoir des données
                TimeBlockPage timeBlockPage = homePage.GoToFlight_TimeBlockPage();
                timeBlockPage.ResetFilters();
                //Effectuer un filtre Search
                timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNo);
                timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);

                //schedule workshops
                timeBlockPage.Schedule();
                timeBlockPage.SetFirstWorkshopParams(j1.ToString("dd/MM/yyyy"), "04:00", j2.ToString("dd/MM/yyyy"), "02:00");

                bool VerifRefreshPage = timeBlockPage.VerifyPageRefresh();
                //Les données sont filtrées en fonction du filtre appliquer
                Assert.IsTrue(VerifRefreshPage, "Le changement des données s'enregistrent mais La Page a été refraichie");
            }
            finally
            {
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNo);
                flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
                flightPage.MassiveDeleteMenus(flightNo, airportFrom, null);

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_ResetAll_Workshop()
        {
            // Arrange: Se connecter en tant qu'admin et accéder à la page TimeBlock
            var homePage = LogInAsAdmin();
            var timeBlockPage = homePage.GoToFlight_TimeBlockPage();

            timeBlockPage.Schedule();

            BulkChangeModal bulkChangemodal = timeBlockPage.BulkChange();

            bulkChangemodal.Fill("5", "1600AM", "2", "1300PM", "2", "1400AM", "1", "1500PM", true);

            bulkChangemodal.Submit();
            var initialDayCorte = timeBlockPage.GetDayCorte();
            var initiatlHourCorte = timeBlockPage.GetHourCorte();
            var initialDayCorteEnsamblaje = timeBlockPage.GetDayCorteEnsamblaje();
            var initiatlHourCorteEnsamblaje = timeBlockPage.GetHourCorteEnsamblaje();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Workshops, "Ensamblaje");
            timeBlockPage.ResetAll();
            timeBlockPage.ResetFilters();
            var finalDayCorte = timeBlockPage.GetDayCorte();
            var finalHourCorte = timeBlockPage.GetHourCorte();
            var finalDayCorteEnsamblaje = timeBlockPage.GetDayCorteEnsamblaje();
            var finalHourCorteEnsamblaje = timeBlockPage.GetHourCorteEnsamblaje();
            Assert.AreEqual(initialDayCorte, finalDayCorte, "rest all non fonctionnel ");
            Assert.AreEqual(finalHourCorte, finalHourCorte, "rest all non fonctionnel ");
            Assert.AreNotEqual(initialDayCorteEnsamblaje, finalDayCorteEnsamblaje, "rest all non fonctionnel ");
            Assert.AreNotEqual(initiatlHourCorteEnsamblaje, finalHourCorteEnsamblaje, "rest all non fonctionnel ");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Filter_DateFrom()
        {
            //Prepare
            string airportFrom = TestContext.Properties["Site"].ToString();
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNo = new Random().Next().ToString();

            // Arrange
            PageObjects.HomePage homePage = LogInAsAdmin();


            //Act
            //Créer un Flight
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            try
            {
                FlightCreateModalPage create = flightPage.FlightCreatePage();
                create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo, null, "00", "23", null, DateUtils.Now);

                //Etre sur l'index des Time Block et avoir des données
                TimeBlockPage timeBlockPage = homePage.GoToFlight_TimeBlockPage();
                timeBlockPage.ResetFilters();
                timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNo);

                //Appliquer filtre Date From 
                // hier
                timeBlockPage.Filter(TimeBlockPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
                timeBlockPage.CloseDatePicker();
                bool filtred = timeBlockPage.CheckTotalNumber() == 1;
                string flightFromList = timeBlockPage.GetFirstFlightNumber();
                Assert.IsTrue(filtred, "Les données ne sont pas filtrées en fonction du filtre appliquer");
                Assert.AreEqual(flightNo, flightFromList, "Le numéro du flight ajouté ne correspond pas");
                // aujourd'hui
                timeBlockPage.Filter(TimeBlockPage.FilterType.DateFrom, DateUtils.Now);
                timeBlockPage.CloseDatePicker();
                filtred = timeBlockPage.CheckTotalNumber() == 1;
                flightFromList = timeBlockPage.GetFirstFlightNumber();
                Assert.IsTrue(filtred, "Les données ne sont pas filtrées en fonction du filtre appliquer");
                Assert.AreEqual(flightNo, flightFromList, "Le numéro du flight ajouté ne correspond pas");
                // demain
                timeBlockPage.Filter(TimeBlockPage.FilterType.DateFrom, DateUtils.Now.AddDays(1));
                filtred = timeBlockPage.CheckTotalNumber() == 0;
                Assert.IsTrue(filtred, "Les données ne sont pas filtrées en fonction du filtre appliquer");
            }
            finally
            {
                homePage.GoToFlights_FlightPage();
                flightPage.MassiveDeleteMenus(flightNo, airportFrom, null);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_TB_Filter_DateTo()
        {
            //Prepare
            string airportFrom = TestContext.Properties["Site"].ToString();
            string airportTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string flightNo = new Random().Next().ToString();

            // Arrange
            PageObjects.HomePage homePage = LogInAsAdmin();

            //Act
            //Créer un Flight
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, airportFrom);
            try
            {
                FlightCreateModalPage create = flightPage.FlightCreatePage();
                create.FillField_CreatNewFlight(flightNo, customer, aircraft, airportFrom, airportTo, null, "00", "23", null, DateUtils.Now);

                //Etre sur l'index des Time Block et avoir des données
                TimeBlockPage timeBlockPage = homePage.GoToFlight_TimeBlockPage();
                timeBlockPage.ResetFilters();
                timeBlockPage.Filter(TimeBlockPage.FilterType.Site, airportFrom);
                timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNo);

                //Appliquer filtre Date To 
                // hier
                timeBlockPage.Filter(TimeBlockPage.FilterType.DateTo, DateUtils.Now.AddDays(1));
                string flightFromList = timeBlockPage.GetFirstFlightNumber();
                bool filtred = timeBlockPage.CheckTotalNumber() == 1;
                Assert.IsTrue(filtred, "Les données ne sont pas filtrées en fonction du filtre appliquer");
                Assert.AreEqual(flightNo, flightFromList, "Le numéro du flight ajouté ne correspond pas");
                // aujourd'hui
                timeBlockPage.Filter(TimeBlockPage.FilterType.DateTo, DateUtils.Now);
                flightFromList = timeBlockPage.GetFirstFlightNumber();
                filtred = timeBlockPage.CheckTotalNumber() == 1;
                Assert.IsTrue(filtred, "Les données ne sont pas filtrées en fonction du filtre appliquer");
                Assert.AreEqual(flightNo, flightFromList, "Le numéro du flight ajouté ne correspond pas");
                // demain
                timeBlockPage.Filter(TimeBlockPage.FilterType.DateTo, DateUtils.Now.AddDays(-1));
                filtred = timeBlockPage.CheckTotalNumber() == 0;
                Assert.IsTrue(filtred, "Les données ne sont pas filtrées en fonction du filtre appliquer");
            }
            finally
            {
                homePage.GoToFlights_FlightPage();
                flightPage.MassiveDeleteMenus(flightNo, airportFrom, null);
            }
        }
    }
}