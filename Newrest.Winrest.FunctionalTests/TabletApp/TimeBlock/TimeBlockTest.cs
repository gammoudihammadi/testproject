using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Tablet;
using Newrest.Winrest.FunctionalTests.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.TimeBlock;
using Newrest.Winrest.FunctionalTests.PageObjects.TabletApp;
using OpenQA.Selenium;

namespace Newrest.Winrest.FunctionalTests.TabletApp.TimeBlock
{
    [TestClass]
    public class TimeBlockTest : TestBase
    {
        private const int _timeout = 600000;


        [TestMethod]
        [Timeout(_timeout)]
        public void TB_TIBL_TimeBlockDisplay()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            var tabletApp = homePage.GotoTabletApp();
            var tabletAppTimeBlock = tabletApp.GotoTabletApp_TimeBlock();
            var isTimeBlockDisplayed = tabletAppTimeBlock.IsTimeBlockDisplayed();
            Assert.IsTrue(isTimeBlockDisplayed, "L'affichage du Time Block a échoué ou un timeout s'est produit.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void TB_TIBL_ValiderHeure_Debut_Fin()
        {
            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            HomePage homePage = LogInAsAdmin();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            FlightCreateModalPage flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);
            TimeBlockPage timeBlockPage = flightPage.GoToFlight_TimeBlockPage();
            timeBlockPage.Filter(TimeBlockPage.FilterType.Search, flightNumber);
            timeBlockPage.AddTimeBlock();
            homePage.Navigate();
            TabletAppPage tabletApp = homePage.GotoTabletApp();
            TimeBlockTabletAppPage tabletAppTimeBlock = tabletApp.GotoTabletApp_TimeBlock();
            tabletAppTimeBlock.Filter(TimeBlockTabletAppPage.FilterType.Search, flightNumber);
            bool isPopUpshowHistoryVisible = tabletAppTimeBlock.IsPopUpshowHistoryVisible();
            Assert.IsTrue(isPopUpshowHistoryVisible, "L'affichage du historique a échoué ou un bug s'est produit.");

        }

    }


}
