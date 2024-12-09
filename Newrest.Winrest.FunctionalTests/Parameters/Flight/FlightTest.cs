using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.Parameters.Flight
{
    [TestClass]
    public class FlightTests : TestBase
    {
        private const int _timeout = 60 * 2 * 10000;

        [TestMethod]
        [Timeout(_timeout)]
        public void SE_FLIG_EditRegistration()
        {
            //Prepare
            string aircraft = "B777";
            string customerLp = TestContext.Properties["CustomerLP"].ToString();
            string customer = "TAG AVIATION ASIA LTD";

            // Arrange
            var homePage = LogInAsAdmin();
            //Act

            var settingsFlightsPage = homePage.GoToParametres_Flights();
            settingsFlightsPage.GoToFlightsRegistrationTypes();
            settingsFlightsPage.ClickOnEdit();
            settingsFlightsPage.WaitPageLoading();
            settingsFlightsPage.EditRegistrationTypeswithCustomer(customerLp);
            settingsFlightsPage.WaitPageLoading();
            settingsFlightsPage.WaitPageLoading();
            settingsFlightsPage.save();
            settingsFlightsPage.WaitPageLoading();

            settingsFlightsPage.ClickOnEdit();
            settingsFlightsPage.EditRegistrationTypes(customer, aircraft);
            settingsFlightsPage.WaitPageLoading();
            settingsFlightsPage.WaitPageLoading();
            settingsFlightsPage.save();
            settingsFlightsPage.WaitPageLoading();

            var customerFromTable = settingsFlightsPage.GetCustomerFromTable();
            var airCraftFromTable = settingsFlightsPage.GetAirCraftFromTable();

            //Assert
            Assert.AreEqual(customerFromTable, customer, "le cutomer n'a pa été modifié ");
            Assert.AreEqual(airCraftFromTable, aircraft, "aircraft  n'a pa été modifié ");
        }

    }
}