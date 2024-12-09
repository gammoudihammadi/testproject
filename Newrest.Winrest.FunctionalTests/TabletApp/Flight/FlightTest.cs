using DocumentFormat.OpenXml.VariantTypes;
using FluentAssertions.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Tablet;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.TabletApp;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Security.Policy;
using System.Web;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;

namespace Newrest.Winrest.FunctionalTests.TabletApp
{
    [TestClass]
    public class FlightTest : TestBase
    {
        private const int Timeout = 600000;
        string flightNumberToday = "";
        string datasheetName = "DatasheetToday" + "-" + DateTime.Now + new Random(1000).Next().ToString();
        string serviceName = "ServiceToday" + DateTime.Now + new Random(1000).Next().ToString();
        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(TB_FLIGHT_EditFlight_SetQuantity):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_DetailAddNotifications):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_Filter_ShowDeliveryNoteComment):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_Filter_InternalFlightRemarks):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_EditFlight_UploadImage):
                    TB_FLIGHT_TestInitialize();
                    break;
                //case nameof(TB_FLIGHT_DetailModifyNotifications):
                //    TB_FLIGHT_TestInitialize();
                //    break;
                //case nameof(TB_FLIGHT_DetailValidateNotifications):
                //    TB_FLIGHT_TestInitialize();
                //    break;
                case nameof(TB_FLIGHT_DetailsFilters_Notification):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_DetailChangeFlight):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_EditFlight_StartAllServiceName):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_EditFlight_StartServiceName):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_EditFlight_TakeAPhoto):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_EditFlight_DoneAllServiceName):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_EditFlight_DoneServiceName):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_EditFlight_DeleteImage):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_DetailColorNotifications):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_EditFlight_FoldUnfold):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_EditFlight_AddGuest):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_Filter_Back_Reset):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_Filter_Drivers):
                    TB_FLIGHT_TestInitialize();
                    break;
                case nameof(TB_FLIGHT_ShowDatasheet):
                    TB_FLIGHT_CreateFlight_WithDatasheet();
                    break;
                case nameof(TB_FLIGHT_StateChangement):
                    TB_FLIGHT_CreateFlight_WithDatasheet();
                    break;
                default:
                    break;
            }
        }

        public void TB_FLIGHT_CreateFlight_WithDatasheet()
        {

            //Prepare recipe
            string ingredient = "Item" + DateTime.Now+ new Random(1000).Next().ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            //Prepare ingredient
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string supplierRef = "10";
            //ingredient + "_SupplierRef";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Prepare
            string guestType = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            flightNumberToday = new Random().Next().ToString();

            string editedGuestTypeQty = "10";
            string code = new Random().Next().ToString();
            DateTime fromDate = DateTime.Today;
            DateTime toDate = DateTime.Today.AddMonths(1);
            string serviceProduction = GenerateName(4);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Create Item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, supplierRef);

            //Create datasheet + Recipe
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            var datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName, ingredient);
            datasheetDetailPage.WaitLoading();

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

            //Create flight
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumberToday, customer, aircraft, site, siteTo);

            //Edit the Guesttype quantity
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberToday);
            var flightEditPage = flightPage.EditFirstFlight(flightNumberToday);
            flightEditPage.AddGuestType();
            flightEditPage.AddService(serviceName);

            flightEditPage.EditGuestTypeQT(editedGuestTypeQty);
            bool datasheetExists = flightEditPage.VerifyDatasheetExist();
            flightPage = flightEditPage.CloseViewDetails();

        }

        [TestCleanup]
        public override void TestCleanup()
        {
            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(TB_FLIGHT_EditFlight_SetQuantity):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_DetailAddNotifications):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_Filter_ShowDeliveryNoteComment):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_Filter_InternalFlightRemarks):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_EditFlight_UploadImage):
                    TB_FLIGHT_TestCleanUp();
                    break;
                //case nameof(TB_FLIGHT_DetailModifyNotifications):
                //    TB_FLIGHT_TestCleanUp();
                //    break;
                //case nameof(TB_FLIGHT_DetailValidateNotifications):
                //    TB_FLIGHT_TestCleanUp();
                //    break;
                case nameof(TB_FLIGHT_DetailsFilters_Notification):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_DetailChangeFlight):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_EditFlight_StartAllServiceName):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_EditFlight_StartServiceName):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_EditFlight_TakeAPhoto):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_EditFlight_DoneAllServiceName):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_EditFlight_DoneServiceName):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_EditFlight_DeleteImage):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_DetailColorNotifications):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_EditFlight_FoldUnfold):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_EditFlight_AddGuest):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_Filter_Back_Reset):
                    TB_FLIGHT_TestCleanUp();
                    break;
                case nameof(TB_FLIGHT_Filter_Drivers):
                    TB_FLIGHT_TestCleanUp();
                    break;

                default:
                    break;
            }
            base.TestCleanup();
        }

        public void TB_FLIGHT_TestInitialize()
        {
            // Prepare
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            string customer1 = TestContext.Properties["CustomerLPFlight"].ToString();
            string customer2 = TestContext.Properties["Bob_Customer"].ToString();
            string BCN = TestContext.Properties["SiteToFlight"].ToString();
            string ACE = TestContext.Properties["SiteLP"].ToString();
            string flightNo = "FlightACE_";
            string serviceName1 = TestContext.Properties["FlightService"].ToString();
            string serviceName2 = TestContext.Properties["ServiceForBOBLPCart"].ToString();
            string guestName = TestContext.Properties["Guest"].ToString();


            string flight1 = flightNo + DateUtils.Now.AddDays(1).ToString("yyyy-MM-dd") + "_ForTabletApp_" + customer1;
            string flight2 = flightNo + DateUtils.Now.AddDays(1).ToString("yyyy-MM-dd") + "_ForTabletApp_" + customer2;

            // Flight 1
            TestContext.Properties[string.Format("FlightDeliveriesId1")] = InsertFlightDeliveries(ACE, customer1, flight1);
            InsertFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId1")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, ACE, BCN);
            TestContext.Properties[string.Format("FlightLegsId1")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId1")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId1")] = InsertFlightLegToGuestTypes(guestName, (int)TestContext.Properties[string.Format("FlightLegsId1")]);
            TestContext.Properties[string.Format("FlightDetailsId1")] = InsertFlightDetailsWithService(serviceName1, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId1")]);
            TestContext.Properties[string.Format("FlightDetailsId3")] = InsertFlightDetailsWithService(serviceName2, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId1")]);

            // Flight 2
            TestContext.Properties[string.Format("FlightDeliveriesId2")] = InsertFlightDeliveries(ACE, customer2, flight2);
            InsertFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId2")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, ACE, BCN);
            TestContext.Properties[string.Format("FlightLegsId2")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId2")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId2")] = InsertFlightLegToGuestTypes(guestName, (int)TestContext.Properties[string.Format("FlightLegsId2")]);
            TestContext.Properties[string.Format("FlightDetailsId2")] = InsertFlightDetailsWithService(serviceName1, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId2")]);
            TestContext.Properties[string.Format("FlightDetailsId4")] = InsertFlightDetailsWithService(serviceName2, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId2")]);

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);

            homePage.Navigate();
            ParametersTablet parametersTablet = homePage.GoToParametres_Tablet();
            parametersTablet.ClickFlightTypeTab();
            parametersTablet.EditFlightTypeColor("Fofo", "Orange", ACE);
            parametersTablet.EditFlightTypeColor("Fifi", "Red", ACE);
            FlightPage flightPage = homePage.GoToFlights_FlightPage();

            flightPage.Filter(FlightPage.FilterType.Sites, ACE);

            // Flight 1
            flightPage.ResetFilter();
            flightPage.SetDateState(DateUtils.Now.AddDays(1));
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1);
            flightPage.SetFirstDriver("Toto");
            SetTabletFlightType("Fofo");
            flightPage = homePage.GoToFlights_FlightPage();

            // Flight 2
            flightPage.ResetFilter();
            flightPage.SetDateState(DateUtils.Now.AddDays(1));
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flight2);
            flightPage.SetFirstDriver("Titi");
            SetTabletFlightType("Fifi");
        }

        private void TB_FLIGHT_TestCleanUp()
        {
            #region deleteFlightDetails
            List<int> FlightDetails = new List<int>();

            for (int i = 1; i <= 4; i++)
            {
                if (TestContext.Properties.Contains($"FlightDetailsId{i}"))
                {
                    FlightDetails.Add((int)TestContext.Properties[$"FlightDetailsId{i}"]);
                }
            }
            DeleteFlightDetails(FlightDetails);
            #endregion

            #region deleteFlightLegToGuestTypes
            List<int> FlightLegToGuestTypes = new List<int>();

            for (int i = 1; i <= 2; i++)
            {
                if (TestContext.Properties.Contains($"FlightLegToGuestTypesId{i}"))
                {
                    FlightLegToGuestTypes.Add((int)TestContext.Properties[$"FlightLegToGuestTypesId{i}"]);
                }
            }
            DeleteFlightLegToGuestTypes(FlightLegToGuestTypes);
            #endregion

            #region deleteFlightLegs
            List<int> FlightLegs = new List<int>();

            for (int i = 1; i <= 2; i++)
            {
                if (TestContext.Properties.Contains($"FlightLegsId{i}"))
                {
                    FlightLegs.Add((int)TestContext.Properties[$"FlightLegsId{i}"]);
                }
            }
            DeleteFlightLegs(FlightLegs);
            #endregion

            #region deleteFlightProperties
            List<int> FlightForProperties = new List<int>();

            for (int i = 1; i <= 2; i++)
            {
                if (TestContext.Properties.Contains($"FlightDeliveriesId{i}"))
                {
                    FlightForProperties.Add((int)TestContext.Properties[$"FlightDeliveriesId{i}"]);
                }
            }
            DeleteFlightProperties(FlightForProperties);
            #endregion

            #region deleteFlightNotificationStatusChange
            List<int> FlightNotificationStatusChanges = new List<int>();

            for (int i = 1; i <= 2; i++)
            {
                if (TestContext.Properties.Contains($"FlightDeliveriesId{i}"))
                {
                    FlightNotificationStatusChanges.Add((int)TestContext.Properties[$"FlightDeliveriesId{i}"]);
                }
            }
            DeleteFlightNotificationStatusChanges(FlightNotificationStatusChanges);
            #endregion
            #region deleteFlightHistories
            List<int> FlightForHistories = new List<int>();

            for (int i = 1; i <= 12; i++)
            {
                if (TestContext.Properties.Contains($"FlightDeliveriesId{i}"))
                {
                    FlightForHistories.Add((int)TestContext.Properties[$"FlightDeliveriesId{i}"]);
                }
            }
            DeleteFlightHistories(FlightForHistories);
            #endregion
            #region deleteFlights
            List<int> Flights = new List<int>();

            for (int i = 1; i <= 2; i++)
            {
                if (TestContext.Properties.Contains($"FlightDeliveriesId{i}"))
                {
                    Flights.Add((int)TestContext.Properties[$"FlightDeliveriesId{i}"]);
                }
            }
            DeleteFlights(Flights);
            #endregion

            #region deleteFlightDeliveries
            List<int> FlightDeliveries = new List<int>();

            for (int i = 1; i <= 2; i++)
            {
                if (TestContext.Properties.Contains($"FlightDeliveriesId{i}"))
                {
                    FlightDeliveries.Add((int)TestContext.Properties[$"FlightDeliveriesId{i}"]);
                }
            }
            DeleteFlightDeliveries(FlightDeliveries);
            #endregion
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_AddComment()
        {
            // Prepare
            string CustomerJob = TestContext.Properties["CustomerJob"].ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string DocSite = TestContext.Properties["SiteLP"].ToString();
            var rnd = new Random();
            string flightNo = "FlightACE_";
            string flight1 = flightNo + "aujourdhui_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_ForTabletApp_" + rnd.Next(); ;
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string text1 = "test test " + rnd.Next();
            string text2 = "test " + rnd.Next();

            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flight1, CustomerJob, aircraft, siteFrom, siteTo, null, "00", "23", null, DateUtils.Now.AddDays(1));
                homePage.Navigate();

                TabletAppPage tabletAppPage = homePage.GotoTabletApp();
                tabletAppPage.Select("mat-select-0", DocSite);
                FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
                tabletFlightPage.ClickSurChooseAViewDefault();
                tabletFlightPage.WaitPageLoading();
                //1. Cliquer sur l'icone filtre
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Search, flight1);
                //2 Sélectionner Filter Show Flights_WithAlertsOnly
                tabletFlightPage.CliqueSurOKFilterBtn();
                DateTime flightDate = tabletFlightPage.GetFirstDateFlight();

                tabletFlightPage.WaitPageLoading();
                tabletFlightPage.SetCommentForFlight(flight1, text1, text2);

                homePage.Navigate();
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                var flightDetail = flightPage.EditFirstFlight(flight1);
                string comment1Flight;
                string comment2Flight;

                flightDetail.GetComments(out comment1Flight, out comment2Flight);
                Assert.AreEqual(text1, comment1Flight);
                Assert.AreEqual(text2, comment2Flight);
            }
            finally
            {
                homePage.Navigate();
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.SetNewState("C");
                flightPage.MassiveDeleteMenus(flight1, siteFrom, null, true);
                flightPage.ResetFilter();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowDock()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Select("mat-select-0", DocSite);

            //Etre sur le module Flight Tablet et avoir des vols
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();

            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowDock, true);
            tabletFlightPage.CliqueSurOKFilterBtn();
            //Verifier la colonne Dock ajoutée ou pas
            var verifyColonneAjoutee1 = tabletFlightPage.VerifyColonneAjoutee("Dock");
            Assert.IsTrue(verifyColonneAjoutee1, "Filter ShowDock ne fonctionne pas.");

            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowDock, false);
            tabletFlightPage.CliqueSurOKFilterBtn();
            var verifyColonneAjoutee2 = tabletFlightPage.VerifyColonneAjoutee("Dock");
            Assert.IsFalse(verifyColonneAjoutee2, "Filter ShowDock ne fonctionne pas.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowDriver()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Select("mat-select-0", DocSite);

            //Etre sur le module Flight Tablet et avoir des vols
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();

            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowDriver, true);
            tabletFlightPage.CliqueSurOKFilterBtn();
            //Verifier la colonne Driver ajoutée ou pas
            var verifyColonneAjoutee1 = tabletFlightPage.VerifyColonneAjoutee("Driver");
            Assert.IsTrue(verifyColonneAjoutee1, "Filter ShowDriver ne fonctionne pas.");
            tabletFlightPage.WaitPageLoading();
            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowDriver, false);
            tabletFlightPage.CliqueSurOKFilterBtn();
            var verifyColonneAjoutee2 = tabletFlightPage.VerifyColonneAjoutee("Driver");
            Assert.IsFalse(verifyColonneAjoutee2, "Filter ShowDriver ne fonctionne pas.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowETA()
        {
            // Prepare
            string DocSite = TestContext.Properties["SiteACE"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Select("mat-select-0", DocSite);

            //Etre sur le module Flight Tablet et avoir des vols
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();
            var dataExist = tabletFlightPage.VerifyDataExist();
            if (dataExist == false)
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            }

            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowETA, true);
            tabletFlightPage.CliqueSurOKFilterBtn();
            //Verifier la colonne ETA ajoutée ou pas
            var verifyColonneAjoutee1 = tabletFlightPage.VerifyColonneAjoutee("ETA");
            Assert.IsTrue(verifyColonneAjoutee1, "Filter ShowETA ne fonctionne pas.");
            tabletFlightPage.WaitPageLoading();
            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowETA, false);
            tabletFlightPage.CliqueSurOKFilterBtn();

            var verifyColonneAjoutee2 = tabletFlightPage.VerifyColonneAjoutee("ETA");
            Assert.IsFalse(verifyColonneAjoutee2, "Filter ShowETA ne fonctionne pas.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowGuest()
        {
            // Prepare
            string DocSite = TestContext.Properties["SiteACE"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Select("mat-select-0", DocSite);

            //Etre sur le module Flight Tablet et avoir des vols
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitLoading();
            var dataExist = tabletFlightPage.VerifyDataExist();
            if (dataExist == false)
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            }
            var numberColumnsBeforeFilter = tabletFlightPage.GetNumberOfColumns();

            //Filter
            tabletFlightPage.WaitLoading();
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowGuest, true);
            tabletFlightPage.CliqueSurOKFilterBtn();
            //Verifier la colonne Guest ajoutée ou pas
            var numberColumnsAfterFilter1 = tabletFlightPage.GetNumberOfColumns();
            var isTrue = numberColumnsAfterFilter1 > numberColumnsBeforeFilter;
            Assert.IsTrue(isTrue, "Filter ShowGuest ne fonctionne pas.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowLeg()
        {
            // Prepare
            string CustomerJob = TestContext.Properties["CustomerJob"].ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string DocSite = TestContext.Properties["SiteLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, CustomerJob, aircraft, siteFrom, siteTo, null, "00", "23", null, DateUtils.Now.AddDays(1));
                WebDriver.Navigate().Refresh();
                homePage.Navigate();

                TabletAppPage tabletAppPage = homePage.GotoTabletApp();
                tabletAppPage.Select("mat-select-0", DocSite);
                //Etre sur le module Flight Tablet et avoir des vols
                FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
                tabletFlightPage.ClickSurChooseAViewDefault();
                tabletFlightPage.WaitLoading();
                //Filter
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowLeg, true);
                tabletFlightPage.CliqueSurOKFilterBtn();
                //Verifier la colonne Leg ajoutée ou pas
                var verifyColonneAjoutee1 = tabletFlightPage.VerifyColonneAjoutee("Legs");
                Assert.IsTrue(verifyColonneAjoutee1, "Filter ShowLeg ne fonctionne pas.");
                tabletFlightPage.WaitPageLoading();
                //Filter
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowLeg, false);
                tabletFlightPage.CliqueSurOKFilterBtn();

                var verifyColonneAjoutee2 = tabletFlightPage.VerifyColonneAjoutee("Legs");
                Assert.IsFalse(verifyColonneAjoutee2, "Filter ShowLeg ne fonctionne pas.");
            }
            finally
            {
                homePage.Navigate();
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.SetNewState("C");
                flightPage.MassiveDeleteMenus(flightNumber, siteFrom, null, true);
                flightPage.ResetFilter();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowSpecMeal()
        {
            // Prepare
            string DocSite = TestContext.Properties["SiteACE"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Select("mat-select-0", DocSite);

            //Etre sur le module Flight Tablet et avoir des vols
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowSpecMeal, true);
            tabletFlightPage.CliqueSurOKFilterBtn();
            //Verifier la colonne SpecMeal ajoutée ou pas

            //Assert
            Assert.IsTrue(tabletFlightPage.VerifyColonneAjoutee("Spec. Meal"), "Filter ShowSpecMeal ne fonctionne pas.");
            tabletFlightPage.WaitPageLoading();
            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowSpecMeal, false);
            tabletFlightPage.CliqueSurOKFilterBtn();

            //Assert
            Assert.IsFalse(tabletFlightPage.VerifyColonneAjoutee("Spec. Meal"), "Filter ShowSpecMeal ne fonctionne pas.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowWorkshop()
        {
            // Prepare
            string DocSite = TestContext.Properties["SiteACE"].ToString();
            string history = "History";
            string corte = "Corte";
            string ensambl = "Ensambl";
            //int nbColumns = 3 + 1; // Show Workshop rajoute une autre colonne "History"
            string unSelectAll = "UnSelect All";
            string selectAll = "Select All";
            //Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Select("mat-select-0", DocSite);

            //Etre sur settings pour cocher juste trois workshops
            var applicationSettings = homePage.GoToApplicationSettings();
            var parametersProduction = applicationSettings.GoToParameters_ProductionPage();
            parametersProduction.WorkshopTab();
            //verify ensamblaje et corte existent
            var exist = parametersProduction.VerifyWorkshopExist("Ensamblaje");
            if (!exist)
            {
                parametersProduction.CreateNewWorkshop(ensambl);
            }
            exist = parametersProduction.VerifyWorkshopExist(corte);
            if (!exist)
            {
                parametersProduction.CreateNewWorkshop(corte);
            }
            exist = parametersProduction.VerifyWorkshopExist(history);
            if (!exist)
            {
                parametersProduction.CreateNewWorkshop(history);
            }

            var parametersTablet = applicationSettings.GoToParametres_Tablet();
            parametersTablet.WorkshopsTab();
            parametersTablet.ShowOnTabletByLabel(ensambl, corte, history);
            parametersTablet.DesactiverTabletByLabel(ensambl, corte, history);

            //Etre sur le module Flight Tablet et avoir des vols
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletAppPage.WaitPageLoading();
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, false);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Workshop, unSelectAll);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Workshop, selectAll);
            tabletFlightPage.CliqueSurOKFilterBtn();
           // var totalColumnsBeforFilter = tabletFlightPage.GetNumberOfColumns();
            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, true);
            tabletFlightPage.CliqueSurOKFilterBtn();

            //Assert
            Assert.IsTrue(tabletFlightPage.VerifyAddedColonnes(), "Filter ShowWorkshop ne fonctionne pas.");

            tabletFlightPage.WaitPageLoading();

            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, false);
            tabletFlightPage.CliqueSurOKFilterBtn();

            //Assert

            Assert.IsFalse(tabletFlightPage.VerifyAddedColonnes(), "Filter ShowWorkshop ne fonctionne pas.");
        
        }
        /// <summary>
        /// Delete Methods 

        private void DeleteFlights(List<int> flightsIds)
        {
            foreach (var flightId in flightsIds)
            {
                DeleteFlight(flightId);
            }
        }

        private void DeleteFlightProperties(List<int> flightsIds)
        {
            foreach (var flightId in flightsIds)
            {
                DeleteFlightPropertie(flightId);
            }
        }

        private void DeleteFlightNotificationStatusChanges(List<int> flightsIds)
        {
            foreach (var flightId in flightsIds)
            {
                DeleteFlightNotificationStatusChange(flightId);
            }
        }

        private void DeleteFlightDeliveries(List<int> flightsIds)
        {
            foreach (var flightId in flightsIds)
            {
                DeleteFlightDeliverie(flightId);
            }
        }

        private void DeleteFlightLegs(List<int> flightsIds)
        {
            foreach (var flightId in flightsIds)
            {
                DeleteFlightLeg(flightId);
            }
        }

        private void DeleteFlightLegToGuestTypes(List<int> flightsIds)
        {
            foreach (var flightId in flightsIds)
            {
                DeleteFlightLegToGuestType(flightId);
            }
        }

        private void DeleteFlightDetails(List<int> flightsIds)
        {
            foreach (var flightId in flightsIds)
            {
                DeleteFlightDetailsWithService(flightId);
            }
        }

        private void DeleteFlightHistories(List<int> flightsIds)
        {
            foreach (var flightId in flightsIds)
            {
                DeleteFlightHistorie(flightId);
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_Customer()
        {
            // Prepare
            string DocSite = "ACE";
            string customer = "TVS";

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletAppPage.WaitPageLoading();

            //Act
            //1. Cliquer sur l'icone filtre
            tabletFlightPage.CliqueSurFilterIcone();
            //2 Sélectionner Filter -Auto Scroll- Customer
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Customer, "UnSelect All");
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Customer, customer);
            //3.Cliquer sur Ok
            tabletFlightPage.CliqueSurOKFilterBtn();
            Assert.IsTrue(tabletFlightPage.IsAllCustomer(customer), "Mauvais filtrage par Customer");

            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Customer, "Select All");
            tabletFlightPage.CliqueSurOKFilterBtn();

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_Workshops()
        {
            // Prepare
            string DocSite = TestContext.Properties["SiteACE"].ToString();
            string workshop = "Corte";

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletAppPage.WaitPageLoading();
            try
            {
                //Act

                //1. Cliquer sur l'icone filtre
                tabletFlightPage.CliqueSurFilterIcone();
                //2 Sélectionner Filter Show Workshops
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, true);
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Workshop, "UnSelect All");
                int totalColumnsBeforFilter = tabletFlightPage.GetNumberOfColumns();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Workshop, workshop);
                //3.Cliquer sur Ok
                tabletFlightPage.CliqueSurOKFilterBtn();

                //Assert
                //Les données sont filtrées et la liste correspond à celle sur Winrest
                var totalColumnsAfterFilter = tabletFlightPage.GetNumberOfColumns();
                Assert.AreEqual(totalColumnsAfterFilter, totalColumnsBeforFilter + 1, "Filter Workshop ne fonctionne pas.");
            }
            finally
            {
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, false);
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Workshop, "Select All");
                tabletFlightPage.CliqueSurOKFilterBtn();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_Search()
        {
            // Prepare
            string DocSite = TestContext.Properties["SiteACE"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();
            var dataExist = tabletFlightPage.VerifyDataExist();
            if (dataExist == false)
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            }
            string flight = tabletFlightPage.GetFlightName(1);
            //1. Cliquer sur l'icone filtre
            tabletFlightPage.CliqueSurFilterIcone();
            //2 Sélectionner Filter Search
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Search, flight);
            tabletFlightPage.CliqueSurOKFilterBtn();
            bool VerifySearchFilter = tabletFlightPage.VerifySearchFilterList(flight);
            Assert.IsTrue(VerifySearchFilter, "Search filter ne fonctionne pas.");
        
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_EDTFrom()
        {
            // Prepare
            string EDTFromH = "02";
            string EDTFromM = "00";
            string CustomerJob = TestContext.Properties["CustomerJob"].ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string DocSite = TestContext.Properties["SiteLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, CustomerJob, aircraft, siteFrom, siteTo, null, EDTFromM, EDTFromH, null, DateUtils.Now.AddDays(1));
                WebDriver.Navigate().Refresh();
                homePage.Navigate();

                //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
                TabletAppPage tabletAppPage = homePage.GotoTabletApp();
                tabletAppPage.Select("mat-select-0", DocSite);
                FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
                tabletFlightPage.ClickSurChooseAViewDefault();
                tabletFlightPage.WaitPageLoading();
                var nbreFlightsLinesBeforeFilter = tabletFlightPage.GetNumberFlightLines();
                //1. Cliquer sur l'icone filtre
                tabletFlightPage.CliqueSurFilterIcone();
                //2 Filter
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.EDTFrom, EDTFromH, EDTFromM);
                tabletFlightPage.CliqueSurOKFilterBtn();
                var nbreFlightsLinesAfterFilter = tabletFlightPage.GetNumberFlightLines();
                if (nbreFlightsLinesAfterFilter > 0)
                {
                    bool verifyFilterETDFrom = tabletFlightPage.VerifyFilterETDFrom(EDTFromH, EDTFromM);
                    Assert.IsTrue(verifyFilterETDFrom, "ETDFrom filter ne fonctionne pas.");
                }
                else
                {
                    Assert.AreNotEqual(nbreFlightsLinesAfterFilter, nbreFlightsLinesBeforeFilter, "ETDFrom filter ne fonctionne pas.");
                }
            }
            finally
            {
                homePage.Navigate();
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.SetNewState("C");
                flightPage.MassiveDeleteMenus(flightNumber, siteFrom, null, true);
                flightPage.ResetFilter();
            }  
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowWithAlertsOnly()
        {
            // Prepare
            string alertTypes1 = "Missing Flight LPCart";
            string modalMissingDockNumber = "MissingFlightLPCart";
            string CustomerJob = TestContext.Properties["CustomerJob"].ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string DocSite = TestContext.Properties["SiteLP"].ToString();
            var rnd = new Random();
            string flightNo = "FlightACE_";
            string flight1 = flightNo + "aujourdhui_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_ForTabletApp_" + rnd.Next(); ;
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string alert = "Missing LPCart";

            //Arrange
            var homePage = LogInAsAdmin();
            var applicationSettings = homePage.GoToApplicationSettings();
            var paramFlights = applicationSettings.GoToParametres_Flights();
            paramFlights.GoToFlightAlerts();
            paramFlights.SelectAlertTypes(alertTypes1);
            bool isExistCustomer = paramFlights.IsCustomerInList(CustomerJob);
            if (isExistCustomer == true)
            {
                paramFlights.DeleteCustomerSubscription(modalMissingDockNumber, CustomerJob, siteFrom);
            }
            paramFlights.AddCustomerSubscriptionToAlertTypes(modalMissingDockNumber, CustomerJob, siteFrom);
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flight1, CustomerJob, aircraft, siteFrom, siteTo, null, "00", "23", null, DateUtils.Now.AddDays(1));
                WebDriver.Navigate().Refresh();
                homePage.Navigate();
                //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
                TabletAppPage tabletAppPage = homePage.GotoTabletApp();
                tabletAppPage.Select("mat-select-0", DocSite);
                FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
                tabletFlightPage.ClickSurChooseAViewDefault();
                tabletFlightPage.WaitPageLoading();
                //tabletFlightPage.SetDate(DateUtils.Now.AddDays(2));
                //1. Cliquer sur l'icone filtre
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Search, flight1);
                //2 Sélectionner Filter Show Flights_WithAlertsOnly
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowFlights_WithAlertsOnly, true);
                tabletFlightPage.CliqueSurOKFilterBtn();

                bool exists = tabletFlightPage.VerifyAllFlightWithAlert(flight1, alert);
                Assert.IsTrue(exists, "Show with Alert Only Filter ne fonctionne pas");
            }
            finally
            {
                homePage.Navigate();
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.SetNewState("C");
                flightPage.MassiveDeleteMenus(flight1, siteFrom, null, true);
                flightPage.ResetFilter();
            }
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowFlightNotWithDoneStatus()
        {
            // Prepare
            string DocSite = TestContext.Properties["SiteLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customerBob = TestContext.Properties["Bob_CustomerFilter"].ToString();
            
            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            try
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo, null, "00", "23", null, DateUtils.Now.AddDays(1));
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.SetNewState("V");
                WebDriver.Navigate().Refresh();
                homePage.Navigate();
                //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
                TabletAppPage tabletAppPage = homePage.GotoTabletApp();
                tabletAppPage.Select("mat-select-0", DocSite);
                FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
                tabletFlightPage.ClickSurChooseAViewDefault();
                tabletFlightPage.WaitPageLoading();

                //1. Cliquer sur l'icone filtre
                tabletFlightPage.CliqueSurFilterIcone();
                //2 Sélectionner Filter Show Flights_NotWithDoneStatut
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowFlights_ValidatedOnly, true);
                tabletFlightPage.CliqueSurOKFilterBtn();

                tabletFlightPage.WaitPageLoading();
                var isValid = tabletFlightPage.IsFlightValidatedOnly();
                Assert.IsTrue(isValid ,"Flight n'est pas valid");

            }
            finally
            {
                homePage.Navigate();
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.SetNewState("C");
                flightPage.MassiveDeleteMenus(flightNumber, siteFrom, null, true);
                flightPage.ResetFilter();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_Drivers()
        {
            // Prepare
            string DocSite = "ACE";
            string driver = "Titi"; // "Tata";// "Toto";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            //1.Cliquer sur l'icone filtre
            tabletFlightPage.CliqueSurFilterIcone();
            //2 Sélectionner Filter Drivers
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowDriver, true);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Driver, "Select All");
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Driver, "UnSelect All");
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Driver, driver);
            //3.Cliquer sur Ok
            tabletFlightPage.CliqueSurOKFilterBtn();

            Assert.IsTrue(tabletFlightPage.IsAllDriver(driver), "Mauvais filtrage par Driver");

            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowDriver, false);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Driver, "Select All");
            tabletFlightPage.CliqueSurOKFilterBtn();
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_TabletFlightTypes()
        {
            // Prepare
            string DocSite = "ACE";
            string tabletFlightType = "None";

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            //1.Cliquer sur l'icone filtre
            tabletFlightPage.CliqueSurFilterIcone();
            //2 Sélectionner Filter Tablet Flight Types
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.TabletFlightType, "UnSelect All");
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.TabletFlightType, tabletFlightType);
            //3.Cliquer sur Ok
            tabletFlightPage.CliqueSurOKFilterBtn();
            //Les données sont filtrées et la liste correspond à celle sur Winrest
            //Assert
            Assert.IsTrue(tabletFlightPage.IsAllTabletFlightType(tabletFlightType), "Mauvais filtrage par Driver");
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.TabletFlightType, "Select All");
            tabletFlightPage.CliqueSurOKFilterBtn();
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_EDTTo()
        {
            // Prepare
            string DocSite = "ACE";
            string EDTFromH = "02";
            string EDTFromM = "00";
            string EDTToH = "02";
            string EDTToM = "05";

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();
            var dataExist = tabletFlightPage.VerifyDataExist();
            if (dataExist == false)
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            }
            var nbreFlightsLinesBeforeFilter = tabletFlightPage.GetNumberFlightLines();
            //1. Cliquer sur l'icone filtre
            tabletFlightPage.CliqueSurFilterIcone();
            //2 Filter
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.EDTFrom, EDTFromH, EDTFromM);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.EDTTo, EDTToH, EDTToM);
            tabletFlightPage.CliqueSurOKFilterBtn();
            var nbreFlightsLinesAfterFilter = tabletFlightPage.GetNumberFlightLines();

            if (nbreFlightsLinesAfterFilter > 0)
            {
                bool verifyFilterETDFrom = tabletFlightPage.VerifyFilterETDFrom(EDTFromH, EDTFromM);
                Assert.IsTrue(verifyFilterETDFrom, "ETDFrom filter ne fonctionne pas.");
                bool verifyFilterETDTo = tabletFlightPage.VerifyFilterETDTo(EDTToH, EDTToM);
                Assert.IsTrue(verifyFilterETDTo, "ETDTo filter ne fonctionne pas.");
            }
            else
            {
                Assert.AreNotEqual(nbreFlightsLinesAfterFilter, nbreFlightsLinesBeforeFilter, "ETDTo filter ne fonctionne pas.");
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowWithNotificationOnly()
        {
            string notification = "Notification " + DateTime.Now.Day + " / " + DateTime.Now.Month + " / " + DateTime.Now.Hour + " / " + DateTime.Now.Minute;
            // Prepare
            string CustomerJob = TestContext.Properties["CustomerJob"].ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string DocSite = TestContext.Properties["SiteLP"].ToString();
            var rnd = new Random();
            string flightNo = "FlightACE_";
            string flight1 = flightNo + "aujourdhui_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_ForTabletApp_" + rnd.Next(); ;
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();


            //Arrange
            var homePage = LogInAsAdmin();

            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flight1, CustomerJob, aircraft, siteFrom, siteTo, null, "00", "23", null, DateUtils.Now.AddDays(1));
                var flightDetailsPage = flightPage.EditFirstFlight(flight1);
                flightDetailsPage.AddGuestType();
                flightDetailsPage.AddService(serviceName);
                WebDriver.Navigate().Refresh();
                homePage.Navigate();

                //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
                TabletAppPage tabletAppPage = homePage.GotoTabletApp();
                tabletAppPage.Select("mat-select-0", DocSite);
                FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
                tabletFlightPage.ClickSurChooseAViewDefault();
                tabletFlightPage.WaitPageLoading();

                var nbreFlightsLinesBeforeFilter = tabletFlightPage.GetNumberFlightLines();
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Search, flight1);
                tabletFlightPage.CliqueSurOKFilterBtn();

                DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
                //var flightName = detailFlight.GetFlightName();
                detailFlight.ShowExtendedMenu();
                detailFlight.ClickFilterButton();
                if (!detailFlight.isNotificationFilterSelected())
                {
                    detailFlight.Filter(DetailFlightTabletAppPage.FilterType.Notification, true);
                }
                detailFlight.ClickFilterOKBtn();
                //Cliquer sur l'icone filtre
                detailFlight.WaitPageLoading();
                detailFlight.ClickSurIconeNotification();
                //detailFlight.DeleteNotifications();
                detailFlight.AddNotification(notification);
                detailFlight.WaitPageLoading();
                var isNotificationAdded = detailFlight.VerifyNotificationisAdded(notification);
                Assert.IsTrue(isNotificationAdded, "la notification n'est pas ajoutée.");
                detailFlight.CloseNotification();
                detailFlight.ShowExtendedMenu();
                detailFlight.ClickBackButton();
                //1. Cliquer sur l'icone filtre
                tabletFlightPage.CliqueSurFilterIcone();
                //2 déSélectionner Filter Show With Notification Only
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWithNotificationOnly, true);
                tabletFlightPage.CliqueSurOKFilterBtn();
                var nbreFlightsLinesAfterFilter = tabletFlightPage.GetNumberFlightLines();
                Assert.AreNotEqual(nbreFlightsLinesBeforeFilter, nbreFlightsLinesAfterFilter, "Show With Notification Only Filter ne fonctionne pas");
            }
            finally
            {
                homePage.Navigate();
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.SetNewState("C");
                flightPage.MassiveDeleteMenus(flight1, siteFrom, null, true);
                flightPage.ResetFilter();
            }
            

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_SaveView()
        {
            // Prepare
            string DocSite = "ACE";
            string workshop = "MAD";
            string nameView1 = "View" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;
            string nameView2 = workshop + "View" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();
            var dataExist = tabletFlightPage.VerifyDataExist();
            if (dataExist == false)
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            }

            tabletFlightPage.ClickSurChooseAView();
            bool hasDelete = tabletFlightPage.DeleteAllViews();
            tabletFlightPage.UnclickSurChooseAView();
            try
            {
                //1. Cliquer sur l'icone filtre
                // on a désormais un "Default View" non supprimable
                tabletFlightPage.CliqueSurFilterIcone(); // leaving delete all combo

                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, false);
                //2.Save View
                tabletFlightPage.SaveView(nameView1);
                int totalColumnsBeforFilter = tabletFlightPage.GetNumberOfColumns();

                //3. Cliquer sur l'icone filtre
                tabletFlightPage.CliqueSurFilterIcone(); // leaving Save View
                tabletFlightPage.CliqueSurFilterIcone();
                //4. Sélectionner Filter Show Workshops
                // on a désormais un "Default View" non supprimable
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, true);
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, true);
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Workshop, "UnSelect All");
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Workshop, "Select All");
                //5.Save View
                tabletFlightPage.SaveView(nameView2);
                tabletFlightPage.WaitPageLoading();

                //6. Assert Verifier views ajoutées
                tabletFlightPage.ClickSurChooseAView();
                tabletFlightPage.WaitPageLoading();
                bool isView1Exist = tabletFlightPage.isViewExist(nameView1);
                Assert.IsTrue(isView1Exist, "View n'est pas ajoutée");
                tabletFlightPage.WaitPageLoading();
                bool isView2Exist = tabletFlightPage.isViewExist(nameView2);
                Assert.IsTrue(isView2Exist, "View n'est pas ajoutée");
                //7. Assert Verify view save data
                tabletFlightPage.WaitPageLoading();
                bool VerifyView1 = tabletFlightPage.VerifyViewSaveData(nameView1, workshop, totalColumnsBeforFilter);
                Assert.IsTrue(VerifyView1, "Échec de la vérification de la view");
                tabletFlightPage.ClickSurChooseAView();
                tabletFlightPage.WaitPageLoading();
                bool VerifyView2 = tabletFlightPage.VerifyViewSaveData(nameView2, workshop, totalColumnsBeforFilter);
                Assert.IsTrue(VerifyView2, "Échec de la vérification de la view");
                tabletFlightPage.ClickSurChooseAView();
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, false);
                tabletFlightPage.CliqueSurOKFilterBtn();
            }
            finally
            {
                //8. Delete views added
                tabletFlightPage.ClickSurChooseAView();

                tabletFlightPage.DeleteAllViews();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_DeleteView()
        {
            // Prepare
            string DocSite = "ACE";
            string nameView = "View" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();
            var dataExist = tabletFlightPage.VerifyDataExist();
            if (dataExist == false)
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            }
            tabletFlightPage.ClickSurChooseAView();
            bool hasDelete = tabletFlightPage.DeleteAllViews();
            // fermer la combobox (qui a "View Default" non effaçable)
            tabletFlightPage.CliqueSurFilterIcone();

            //1. Cliquer sur l'icone filtre
            tabletFlightPage.CliqueSurFilterIcone();
            //2.Save and Add View
            tabletFlightPage.SaveView(nameView);

            //3. Verifier view ajoutée ou pas
            tabletFlightPage.ClickSurChooseAView();
            tabletFlightPage.WaitPageLoading();
            var isViewExist = tabletFlightPage.isViewExist(nameView);
            Assert.IsTrue(isViewExist, "View n'est pas ajoutée");

            //4. Delete view ajoutée
            tabletFlightPage.DeleteView(nameView);
            tabletFlightPage.WaitPageLoading();
            var isViewNotExist = tabletFlightPage.isViewExist(nameView);
            Assert.IsFalse(isViewNotExist, "View n'est pas supprimée");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_MultipleSaveView()
        {
            // Prepare
            string DocSite = "ACE";
            string nameView1 = "View 1" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;
            string nameView2 = "View 2" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;
            string nameView3 = "View 3" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;
            int counter = 0;

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();
            var dataExist = tabletFlightPage.VerifyDataExist();
            if (dataExist == false)
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            }
            tabletFlightPage.ClickSurChooseAView();
            bool hasDelete = tabletFlightPage.DeleteAllViews();
            tabletFlightPage.UnclickSurChooseAView();
            // sortir de combobox View ("Default View" non effaçable)
            try
            {
                tabletFlightPage.CliqueSurFilterIcone();

                //1. Cliquer sur l'icone filtre
                tabletFlightPage.CliqueSurFilterIcone();
                //2.Save View
                tabletFlightPage.SaveView(nameView1);
                counter++;
                tabletFlightPage.WaitPageLoading(); // le combobox "Choose a view" clignote un coup
                                                    //1. Cliquer sur l'icone filtre
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.WaitLoading();
                //2.Save View
                tabletFlightPage.SaveView(nameView2);
                counter++;
                tabletFlightPage.WaitPageLoading(); // le combobox "Choose a view" clignote un coup
                                                    //1. Cliquer sur l'icone filtre
                tabletFlightPage.CliqueSurFilterIcone();
                //2.Save View
                tabletFlightPage.SaveView(nameView3);
                counter++;
                tabletFlightPage.WaitPageLoading(); // le combobox "Choose a view" clignote un coup

                tabletFlightPage.ClickSurChooseAView();
                int numberOfViews = tabletFlightPage.checkNumberOfViews();
                Assert.AreEqual(numberOfViews, counter + 1, "Le nombre de Views attendu ne correspond pas au nombre de Views réel.");
            }
            finally
            {
                tabletFlightPage.DeleteAllViews();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowByTabletFlightType()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Etre sur le module Flight dans Tablet et avoir des vols
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowFlights_ByTabletFlightType, true);
            tabletFlightPage.CliqueSurOKFilterBtn();

            List<string> typesVerif = tabletFlightPage.GetFlightTypes();
            List<string> typesApres = tabletFlightPage.GetFlightTypes();
            typesVerif.Sort();
            for (int i = 0; i < typesVerif.Count; i++)
            {
                Assert.AreEqual(typesVerif[i], typesApres[i], "FlightType non par ordre alphabetique");
            }

            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowFlights_ByTabletFlightType, false);
            tabletFlightPage.CliqueSurOKFilterBtn();

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_DetailAddNotifications()
        {
            // Prepare
            string DocSite = "ACE";
            string notification = "Notification" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.ShowExtendedMenu();
            detailFlight.ClickFilterButton();
            if (!detailFlight.isNotificationFilterSelected())
            {
                detailFlight.Filter(DetailFlightTabletAppPage.FilterType.Notification, true);
            }
            detailFlight.ClickFilterOKBtn();
            detailFlight.ClickSurIconeNotificationCloche();
            detailFlight.DeleteNotifications();
            detailFlight.AddNotification(notification);
            detailFlight.WaitPageLoading();
            var isNotificationAdded = detailFlight.VerifyNotificationAdded(notification);
            Assert.IsTrue(isNotificationAdded, "la notification n'est pas ajoutée.");

            detailFlight.DeleteNotifications();
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowDeliveryNoteComment()
        {
            // Prepare
            string DocSite = "ACE";
            Random r = new Random();
            string text = "delivery note " + r.Next();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            //Etre sur le module Flight dans Tablet et avoir des vols
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);

            //Etre sur un flight
            //1. Etre sur un Flight
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();

            //2.Cliquer sur...
            DetailFlightTabletAppPage details = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            int iNoFLight = details.GetNoFlight();
            //3.Cliquer sur Filter
            //4.La liste des filtres s'ouvre
            details.ShowExtendedMenu();
            details.ClickFilterButton();

            //5.Choisir un filtre "show Deliveray Note Comment" et cliquer sur OK
            details.Filter(DetailFlightTabletAppPage.FilterType.ShowDeliveryNoteComment, true);
            details.ClickFilterOKBtn();

            //Les données sont filtrées et le champs Delivery Note Comment
            details.SetDeliveryNoteComment(text);
            //peut etre afficher ou rempli et correspond au commentaire dans l'index des flights
            details.ShowExtendedMenu();
            tabletFlightPage = details.ClickBackButton();
            WebDriver.Navigate().Refresh();
            tabletFlightPage.WaitPageLoading();
            string textDansIndex = tabletFlightPage.GetDeliveryNoteComment(iNoFLight);
            Assert.AreEqual(text, textDansIndex, "problème commentaire Delivery Note");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_InternalFlightRemarks()
        {
            // Prepare
            string DocSite = "ACE";
            Random r = new Random();
            string text = "internal flight remarks " + r.Next();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            //Etre sur le module Flight dans Tablet et avoir des vols
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);

            //Etre sur un flight
            //1. Etre sur un Flight
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();

            //2.Cliquer sur...
            DetailFlightTabletAppPage details = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            int iNoFLight = details.GetNoFlight();
            //3.Cliquer sur Filter
            //4.La liste des filtres s'ouvre
            details.ShowExtendedMenu();
            details.ClickFilterButton();

            //5.Choisir un filtre "show Deliveray Note Comment" et cliquer sur OK
            details.Filter(DetailFlightTabletAppPage.FilterType.ShowInternalFlightRemarks, true);
            details.ClickFilterOKBtn();

            //Les données sont filtrées et le champs Delivery Note Comment
            details.SetInternalFlightRemarks(text);
            details.WaitPageLoading();
            //peut etre afficher ou rempli et correspond au commentaire dans l'index des flights
            details.ShowExtendedMenu();
            tabletFlightPage = details.ClickBackButton();
            WebDriver.Navigate().Refresh();
            tabletFlightPage.WaitPageLoading();
            string textDansIndex = tabletFlightPage.GetInternalFlightRemarks(iNoFLight);
            Assert.AreEqual(text, textDansIndex, "problème commentaire Delivery Note");
        }

        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_EditFlight_UploadImage()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            //Etre sur le module Flight dans Tablet et avoir des vols
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);

            //Etre sur le détail d'un vol et avoir des services Name
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage details = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            //1. Cliquer sur le plateau avec les couverts
            //2.Une pop'up s'ouvre
            PlateauFlightTabletAppPage plateau = details.EditPlateau();
            //3. Upload image
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            plateau.RemoveImage();
            plateau.UploadImage(fiUpload);
            var verifyImage1 = plateau.VerifyImage();
            Assert.IsTrue(verifyImage1, "image non affiché en haut à gauche");
            plateau.RemoveImage();
            var verifyImage2 = plateau.VerifyImage();
            Assert.IsFalse(verifyImage2, "image affiché en haut à gauche");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_DetailModifyNotifications()
        {
            // Prepare
            string CustomerJob = TestContext.Properties["CustomerJob"].ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string DocSite = TestContext.Properties["SiteLP"].ToString();
            var rnd = new Random();
            string flightNo = "FlightACE_";
            string flight1 = flightNo + "aujourdhui_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_ForTabletApp_" + rnd.Next(); ;
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string notification = "Notification " + DateTime.Now.Day + " / " + DateTime.Now.Month + " / " + DateTime.Now.Hour + " / " + DateTime.Now.Minute;
            string notification2 = " -> edit";

            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();

            try
            {
                // Create Flight With Guest Type And Service
                flightPage.ResetFilter();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flight1, CustomerJob, aircraft, siteFrom, siteTo, null, "00", "23", null, DateUtils.Now.AddDays(1));
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1);
                var flightDetailsPage = flightPage.EditFirstFlight(flight1);
                flightDetailsPage.AddGuestType();
                flightDetailsPage.AddService(serviceName);
                WebDriver.Navigate().Refresh();
                homePage.Navigate();

                //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
                TabletAppPage tabletAppPage = homePage.GotoTabletApp();
                tabletAppPage.Select("mat-select-0", DocSite);
                FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
                tabletFlightPage.ClickSurChooseAViewDefault();
                tabletFlightPage.WaitPageLoading();

                // Filter Flight Created
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Search, flight1);
                tabletFlightPage.CliqueSurOKFilterBtn();

                tabletFlightPage.WaitPageLoading();
                DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
                tabletFlightPage.WaitPageLoading();

                detailFlight.ShowExtendedMenu();
                detailFlight.ClickFilterButton();
                if (!detailFlight.isNotificationFilterSelected())
                {
                    detailFlight.Filter(DetailFlightTabletAppPage.FilterType.Notification, true);
                }
                detailFlight.ClickFilterOKBtn();
                detailFlight.ClickSurIconeNotification();
                detailFlight.AddNotification(notification);
                detailFlight.WaitPageLoading();
                var verifyNotificationAdded = detailFlight.VerifyNotificationisAdded(notification);
                Assert.IsTrue(verifyNotificationAdded, "la notification n'est pas ajoutée.");

                detailFlight.EditNotification(notification2);
                detailFlight.ClickSurIconeNotification();
                detailFlight.WaitPageLoading();

                var verifyNotificationEdited = detailFlight.VerifyNotificationEdited(notification);
                Assert.IsTrue(verifyNotificationEdited, "La notification n'est pas modifiée.");
                detailFlight.DeleteNotificationsInServiceFlight();
            }
            finally
            {
                // Filter FLight Created and Delete 
                homePage.Navigate();
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.SetNewState("C");
                flightPage.MassiveDeleteMenus(flight1, siteFrom, null, true);
                flightPage.ResetFilter();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_DetailValidateNotifications()
        {
            // Prepare
            string CustomerJob = TestContext.Properties["CustomerJob"].ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string DocSite = TestContext.Properties["SiteLP"].ToString();
            var rnd = new Random();
            string flightNo = "FlightACE_";
            string flight1 = flightNo + "aujourdhui_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_ForTabletApp_" + rnd.Next(); ;
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string notification = "Notification " + DateTime.Now.Day + " / " + DateTime.Now.Month + " / " + DateTime.Now.Hour + " / " + DateTime.Now.Minute;

            //Arrange
            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                // Create Flight With Guest Type And Service
                flightPage.ResetFilter();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flight1, CustomerJob, aircraft, siteFrom, siteTo, null, "00", "23", null, DateUtils.Now.AddDays(1));
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1);
                var flightDetailsPage = flightPage.EditFirstFlight(flight1);
                flightDetailsPage.AddGuestType();
                flightDetailsPage.AddService(serviceName);
                WebDriver.Navigate().Refresh();
                homePage.Navigate();

                //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
                TabletAppPage tabletAppPage = homePage.GotoTabletApp();
                tabletAppPage.Select("mat-select-0", DocSite);
                FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
                tabletFlightPage.ClickSurChooseAViewDefault();
                tabletFlightPage.WaitPageLoading();
                
                // Filter Flight Created
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Search, flight1);
                tabletFlightPage.CliqueSurOKFilterBtn();

                tabletFlightPage.WaitPageLoading();
                DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
                tabletFlightPage.WaitPageLoading();
                // Verify Filter Notification Selected Or Not
                detailFlight.ShowExtendedMenu();
                detailFlight.ClickFilterButton();
                if (!detailFlight.isNotificationFilterSelected())
                {
                    detailFlight.Filter(DetailFlightTabletAppPage.FilterType.Notification, true);
                }
                // Add Notification
                detailFlight.ClickFilterOKBtn();
                detailFlight.ClickSurIconeNotification();
                detailFlight.AddNotification(notification);
                detailFlight.WaitPageLoading();
                // Verify Notification Added
                bool verifyNotificationAdded = detailFlight.VerifyNotificationisAdded(notification);
                Assert.IsTrue(verifyNotificationAdded, "la notification n'est pas ajoutée.");
                // Validate Notification
                detailFlight.ValideNotification();
                detailFlight.WaitPageLoading();
                detailFlight.ClickSurIconeNotification();
                // Verify Notification Is Validate
                bool verifyNotificationValidee = detailFlight.VerifyNotificationValidee();
                Assert.IsTrue(verifyNotificationValidee, "La notification n'est pas validée.");
                detailFlight.DeleteNotificationsInServiceFlight();
            }
            finally
            {
                // Filter FLight Created and Delete 
                homePage.Navigate();
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.SetNewState("C");
                flightPage.MassiveDeleteMenus(flight1, siteFrom, null, true);
                flightPage.ResetFilter();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_DetailsFilters_Notification()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.WaitPageLoading();
            detailFlight.ShowExtendedMenu();
            detailFlight.ClickFilterButton();
            if (!detailFlight.isNotificationFilterSelected())
            {
                detailFlight.Filter(DetailFlightTabletAppPage.FilterType.Notification, true);
            }
            detailFlight.ClickFilterOKBtn();
            detailFlight.WaitPageLoading();
            var verifyNotificationFiltre = detailFlight.NotifacationIsVisible();
            Assert.IsTrue(verifyNotificationFiltre, "Filter Notification ne fonctionne pas.");
            detailFlight.ShowExtendedMenu();
            detailFlight.ClickFilterButton();
            if (detailFlight.isNotificationFilterSelected())
            {
                detailFlight.Filter(DetailFlightTabletAppPage.FilterType.Notification, false);
            }
            detailFlight.ClickFilterOKBtn();
            var verifyNotification = detailFlight.NotifacationIsVisible();
            detailFlight.WaitPageLoading();
            Assert.IsFalse(verifyNotification, "Filter Notification ne fonctionne pas.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_DetailChangeFlight()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.WaitPageLoading();
            var flightNameExist = detailFlight.GetFlightName();
            detailFlight.ClickListFlight();
            detailFlight.ChangeFlight(flightNameExist);
            var nameFlight = detailFlight.GetFlightName();
            Assert.AreNotEqual(nameFlight, flightNameExist, "Flight n'est pas changée.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_EditFlight_SetQuantity()
        {
            // Prepare
            string DocSite = "ACE";
            string qty = "5";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.WaitPageLoading();
            PlateauFlightTabletAppPage plateauFlightTabletAppPage = detailFlight.EditPlateau();
            plateauFlightTabletAppPage.SetQuantity(qty);
            var isRecalculate = plateauFlightTabletAppPage.isRecalculate(qty);
            Assert.IsTrue(isRecalculate, "Les typed pax ne sont pas recalculés.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_EditFlight_StartAllServiceName()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.WaitPageLoading();
            if (detailFlight.isDoneAll())
            {
                detailFlight.StartAllServices();
                if (!detailFlight.isAlreadyStarted())
                {
                    detailFlight.StartAllServices();
                }
            }
            else
            {
                detailFlight.StartAllServices();
            }
            Assert.IsTrue(detailFlight.isStartedAll(), "Les services ne passent pas Started.");
            detailFlight.WaitPageLoading();
            detailFlight.StartAllServices();
            detailFlight.StartAllServices();

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_EditFlight_StartServiceName()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.WaitPageLoading();

            HandleFlightServiceState(detailFlight);
          
            Assert.IsTrue(detailFlight.isStarted(), "Le bouton n'est pas passé à l'état 'Started' comme attendu.");

        }
        private void HandleFlightServiceState(DetailFlightTabletAppPage detailFlight)
        {
            // Obtenir l'état actuel du bouton
            string buttonText = detailFlight.GetButtonState();
            detailFlight.WaitPageLoading();

            switch (buttonText)
            {
                case "Start":
                    // Si le bouton est "Start", on clique pour passer à "Started"
                    detailFlight.ClickOnStart();
                    detailFlight.WaitPageLoading();

                    break;

                case "Started":
                    detailFlight.ClickOnStarted();
                    detailFlight.WaitPageLoading();
                    detailFlight.ClickOnDone();
                    detailFlight.WaitPageLoading();
                    detailFlight.ClickOnStart();
                    detailFlight.WaitPageLoading();

                    return;

                case "Done":
                    // Si le bouton est "Done", cliquer sur "Done" puis "Start" pour revenir à "Started"
                    detailFlight.ClickOnDone();
                    detailFlight.WaitPageLoading();
                    detailFlight.ClickOnStart();
                    detailFlight.WaitPageLoading();

                    break;

                default:
                    // Si un état inconnu est rencontré, lever une exception
                    throw new InvalidOperationException($"État inconnu du bouton : {buttonText}");
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_EditFlight_TakeAPhoto()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.WaitPageLoading();
            PlateauFlightTabletAppPage plateauFlightTabletAppPage = detailFlight.EditPlateau();
            plateauFlightTabletAppPage.RemoveImage();
            plateauFlightTabletAppPage.TakeAPhoto();
            Assert.IsTrue(plateauFlightTabletAppPage.VerifyImage(), "image non affiché en haut à gauche");
            plateauFlightTabletAppPage.RemoveImage();
            Assert.IsFalse(plateauFlightTabletAppPage.VerifyImage(), "image affiché en haut à gauche");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_EditFlight_DoneAllServiceName()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.WaitPageLoading();
            if (detailFlight.isDone())
            {
                detailFlight.DoneServices();
                detailFlight.WaitLoading();
                detailFlight.StartAllServices();
                detailFlight.WaitLoading();
            }
            if (!detailFlight.isAlreadyStarted())
            {
                detailFlight.StartAllServices();
                detailFlight.WaitLoading();
            }
            detailFlight.StartAllServices();
            detailFlight.WaitLoading();
            Assert.IsTrue(detailFlight.isDoneAll(), "Les services ne passent pas Done.");
            detailFlight.StartAllServices();


        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_EditFlight_DoneServiceName()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);

            var tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            //Vérification des données et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int maxRetry = 5; // Pour limiter  le nombre de tentatives : eviter les boucles infinies
            int offSetDay = 1;

            while (offSetDay <= maxRetry)
            {
                if (dataExist && tabletFlightPage.IsValidFlightWithServices())
                {
                    break; // Les données valides sont trouvées
                }

                // Définir une nouvelle date en reculant d'un jour
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
                tabletFlightPage.WaitPageLoading();
            }

            Assert.IsTrue(offSetDay <= maxRetry, "Les données valides n'ont pas été trouvées après plusieurs tentatives.");

            // Modifier le vol et marquer les services comme "Done"
            var detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.WaitPageLoading();

            HandleFlightServiceStateToDone(detailFlight);

            //Assert
            Assert.IsTrue(detailFlight.isDone(), "Le bouton n'est pas passé à l'état 'Done' comme attendu.");
        }

        private void HandleFlightServiceStateToDone(DetailFlightTabletAppPage detailFlight)
        {
            // Obtenir l'état actuel du bouton
            string buttonText = detailFlight.GetButtonState();
            detailFlight.WaitPageLoading();

            switch (buttonText)
            {
                case "Start":
                    // Si le bouton est "Start", on clique pour passer à "Started"
                    detailFlight.ClickOnStart();
                    detailFlight.WaitPageLoading();
                    detailFlight.ClickOnStarted();
                    detailFlight.WaitPageLoading();

                    break;

                case "Started":
                    detailFlight.ClickOnStarted();
                    detailFlight.WaitPageLoading();

                    break;

                case "Done":
                    // Si le bouton est "Done", cliquer sur "Done" puis "Start" pour revenir à "Started"
                    detailFlight.ClickOnDone();
                    detailFlight.WaitPageLoading();
                    detailFlight.ClickOnStart();
                    detailFlight.WaitPageLoading();
                    detailFlight.ClickOnStarted();
                    detailFlight.WaitPageLoading();

                    break;

                default:
                    // Si un état inconnu est rencontré, lever une exception
                    throw new InvalidOperationException($"État inconnu du bouton : {buttonText}");
            }
        }

        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_EditFlight_DeleteImage()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.WaitPageLoading();
            PlateauFlightTabletAppPage plateauFlightTabletAppPage = detailFlight.EditPlateau();
            if (plateauFlightTabletAppPage.VerifyImage())
            {
                plateauFlightTabletAppPage.RemoveImage();
            }
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            plateauFlightTabletAppPage.UploadImage(fiUpload);
            plateauFlightTabletAppPage.WaitPageLoading();
            bool verifyImage1 = plateauFlightTabletAppPage.VerifyImage();
            Assert.IsTrue(verifyImage1, "image non affiché en haut à gauche");
            plateauFlightTabletAppPage.RemoveImage();
            plateauFlightTabletAppPage.WaitPageLoading();
            bool VerifyImage2 = plateauFlightTabletAppPage.VerifyImage();
            Assert.IsFalse(VerifyImage2, "image affiché en haut à gauche");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_DetailColorNotifications()
        {
            // Prepare
            string DocSite = "ACE";
            Random r = new Random();
            string notification = "Notification" + r.Next();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            tabletFlightPage.WaitPageLoading();

            detailFlight.ShowExtendedMenu();
            detailFlight.ClickFilterButton();
            if (!detailFlight.isNotificationFilterSelected())
            {
                detailFlight.Filter(DetailFlightTabletAppPage.FilterType.Notification, true);
            }
            detailFlight.ClickFilterOKBtn();
            detailFlight.ClickSurClocheIcone();
            detailFlight.DeleteNotifications();
            detailFlight.AddNotification(notification);
            detailFlight.WaitPageLoading();
            detailFlight.ClickSurClocheIcone();
            detailFlight.ActivateNotification();
            detailFlight.WaitPageLoading();
            bool isNotificationActive = detailFlight.VerifyNotificationActive();
            Assert.IsTrue(isNotificationActive, "la notification n'est pas activée.");
            detailFlight.DeleteNotifications();
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_EditFlight_FoldUnfold()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.ClickRecipeButton();
            detailFlight.ClickFoldUnfoldButton();
            bool isRecipeDetailOpen = detailFlight.isRecipeDetailOpen();
            Assert.IsTrue(isRecipeDetailOpen, "Le détail d'une recette n'est pas déplié.");

            detailFlight.ClickFoldUnfoldButton();
            bool isRecipeDetailNotOpen = detailFlight.isRecipeDetailOpen();
            Assert.IsFalse(isRecipeDetailNotOpen, "Le détail d'une recette n'est pas plié.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_EditFlight_AddGuest()
        {
            // Prepare
            string DocSite = "ACE";
            string GestToAdd = "SPML FC";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();

            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();

            var flightName = detailFlight.GetFlightName();
            var flightDate = detailFlight.GetFlightDate();

          
                detailFlight.ShowExtendedMenu();
                detailFlight.ClickAddGuestButton();
                detailFlight.SelectGuestToAdd(GestToAdd);
            detailFlight.WaitPageLoading();
                detailFlight.ValidateAddGuest();

                detailFlight.WaitPageLoading();
                bool guestIsPresent = detailFlight.VerifyGuestIsPresent(GestToAdd);
                Assert.IsTrue(guestIsPresent, "Le guest n'est pas ajouté.");  
          
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_Date()
        {
            // Prepare
            string DocSite = "ACE";
            //Arrange
            var homePage = LogInAsAdmin();
            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
        
            tabletFlightPage.SetDate(DateUtils.Now.AddDays(-1));

            string tabletFlight = tabletFlightPage.GetFirstFlightNumber();
            tabletFlightPage.WaitPageLoading();
            LogInAsAdmin();
            var homePage2 = new HomePage(WebDriver, TestContext);
            homePage2.Navigate();
            //Aller au module flight de winrest
            FlightPage flightPage = homePage2.GoToFlights_FlightPage();
            flightPage.WaitPageLoading();
            flightPage.ResetFilter();
            flightPage.SetSiteFilter(DocSite);
            flightPage.WaitPageLoading();
            flightPage.SetDateState(DateUtils.Now);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, tabletFlight);
            var flightName = flightPage.GetFirstFlightNumber();
            Assert.AreEqual(tabletFlight, flightName, "Les vols J affichés sur tablet ne sont pas les mêmes que sur Winrest à J+1");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowByFlightDate()
        {
            // Prepare
            string DocSite = "ACE";
            string date = DateTime.Now.ToString("dd-MMM", new CultureInfo("fr-FR"));

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowFlights_ByflightDate, "true");

            //IsSortedByDate
            Assert.IsTrue(tabletFlightPage.IsSortedByDate(date), "Mauvais sort par flight date");

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowFlightBarsetsToo()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();

            tabletFlightPage.ClickSurChooseAViewDefault();
            int numberOfFlightsBeforeFilter = tabletFlightPage.GetNumberFlightLines();

            //1. Cliquer sur l'icone filtre
            tabletFlightPage.CliqueSurFilterIcone();
            //2. Sélectionner Show Flights Barsets Too
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowFlights_BarsetsToo, "true");
            tabletFlightPage.WaitPageLoading();

            //3.Cliquer sur Ok
            tabletFlightPage.CliqueSurOKFilterBtn();

            int numberOfFlightsAfterFilter = tabletFlightPage.GetNumberFlightLines();

            Assert.AreEqual(numberOfFlightsBeforeFilter, numberOfFlightsAfterFilter, "Mauvais affichage des barsets");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowByFlightMajorOnly()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();
            var dataExist = tabletFlightPage.VerifyDataExist();
            if (dataExist == false)
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            }
            var defaultResult = tabletFlightPage.GetNumberFlightLines();
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowFlights_MajorOnly, true);
            tabletFlightPage.WaitPageLoading();
            //3.Cliquer sur Ok
            tabletFlightPage.CliqueSurOKFilterBtn();
            var resultAfterFilter = tabletFlightPage.GetNumberFlightLines();
            Assert.AreNotEqual(defaultResult, resultAfterFilter, "Show Flights by Major Only Filter ne fonctionne pas");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_ShowFlightValidatedOnly()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();
            var dataExist = tabletFlightPage.VerifyDataExist();
            Assert.IsTrue(dataExist);
            var nbreFlightsLinesBeforeFilter = tabletFlightPage.GetNumberFlightLines();
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowFlights_ValidatedOnly, true);
            tabletFlightPage.CliqueSurOKFilterBtn();
            var nbreFlightsLinesAfterFilter = tabletFlightPage.GetNumberFlightLines();
            var homePage2 = new HomePage(WebDriver, TestContext);
            homePage2.Navigate();
            FlightPage flightPage = homePage2.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, DocSite);
            flightPage.Filter(FlightPage.FilterType.Status, "Valid");
            flightPage.SetDateState(DateTime.Now);
            var flightPageNumber = flightPage.CheckTotalNumber();
            Assert.AreEqual(flightPageNumber, nbreFlightsLinesAfterFilter, "Show validated Only Filter ne fonctionne pas");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Modification_QtySces_Guest()
        {
            // Prepare
            DateTime DateWithValidFlight = new DateTime(2024, 8, 21, 12, 0, 0);
            string Qty = new Random().Next(1, 30).ToString();


            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur Tablet app et sélectionner un site avec des vols sur Winrest
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            /* set a date with a valid flight that we can update the Qty of */
            tabletFlightPage.SetDate(DateWithValidFlight);
            tabletFlightPage.EditFirstFlight();
            /*Edit flight Qty*/
            tabletFlightPage.EditFlightQTY(Qty);
            /* Edit Guest flight */
            tabletFlightPage.EditGuestQTY(Qty);

            /* refresh */
            WebDriver.Navigate().Refresh();

            /* get the flight & guest qty after refresh */
            string flight_Qty = tabletFlightPage.GetFlightQTY();
            string guest_Qty = tabletFlightPage.GetGuestQTY();

            /* the flight OR the guest Qty must be updated */
            bool is_qty_updated = flight_Qty.IsSameOrEqualTo(Qty) || guest_Qty.IsSameOrEqualTo(Qty);

            /*Assert*/
            Assert.IsTrue(is_qty_updated, "The flight AND the guest qty is not updated ");

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Show_workshop()
        {
            // Prepare
            string CustomerJob = TestContext.Properties["CustomerJob"].ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string DocSite = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            var rnd = new Random();
            string flightNo = "FlightACE_";
            string flight1 = flightNo + "aujourdhui_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_ForTabletApp_" + rnd.Next(); ;
            string flight2 = flightNo + "aujourdhui_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_ForTabletApp_" + rnd.Next(); ;

            var homePage = LogInAsAdmin();
            var flightPage = homePage.GoToFlights_FlightPage();
            try
            {
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flight1, CustomerJob, aircraft, siteFrom, siteTo, null, "00", "23", null, DateUtils.Now.AddDays(2));
                WebDriver.Navigate().Refresh();
                flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flight2, CustomerJob, aircraft, siteFrom, siteTo, null, "00", "23", null, DateUtils.Now.AddDays(2));
                WebDriver.Navigate().Refresh();
                homePage.Navigate();


                TabletAppPage tabletAppPage = homePage.GotoTabletApp();
                tabletAppPage.Select("mat-select-0", DocSite);
                FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
                tabletFlightPage.ClickSurChooseAViewDefault();
                tabletFlightPage.WaitPageLoading();
                tabletFlightPage.SetDate(DateUtils.Now.AddDays(1));
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, true);
                tabletFlightPage.CliqueSurOKFilterBtn();

                var flightNameBeforeDone = tabletFlightPage.GetFlightName(1);
                tabletFlightPage.PutStateToDone();
                var flightNameBeforeStarted = tabletFlightPage.GetFlightName(2);
                tabletFlightPage.PutStateToStarted();
                tabletFlightPage.WaitPageLoading();

                var isDoneWorked = tabletFlightPage.CheckFlightState(flightNameBeforeDone, 1, "DONE");
                var isStartedWorked = tabletFlightPage.CheckFlightState(flightNameBeforeStarted, 2, "STARTED");
                var allButtonsWorked = isDoneWorked && isStartedWorked;

                Assert.IsTrue(allButtonsWorked, "the button text is not changing correctly");
            }
            finally
            {
                homePage.Navigate();
                homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight1);
                flightPage.SetDateState(DateUtils.Now.AddDays(2));
                flightPage.SetNewState("C");
                flightPage.MassiveDeleteMenus(flight1, siteFrom, null, true);
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight2);
                flightPage.SetDateState(DateUtils.Now.AddDays(2));
                flightPage.SetNewState("C");
                flightPage.MassiveDeleteMenus(flight2, siteFrom, null, true);
                flightPage.ResetFilter();
            }
            
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_ShowFlights_WithAlertsOnly()
        {
            var homePage = LogInAsAdmin();
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowFlights_WithAlertsOnly, true);
            tabletFlightPage.CliqueSurOKFilterBtn();
            Assert.IsTrue(tabletFlightPage.CheckAlertExist(), "alert only filter does not work");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Filter_Back_Reset()
        {
            // Prepare
            string DocSite = "ACE";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            tabletFlightPage.WaitPageLoading();

            // Vérification et ajustement de la date si nécessaire
            var dataExist = tabletFlightPage.VerifyDataExist();

            int offSetDay = 1;
            while (!(dataExist && tabletFlightPage.IsValidFlightWithServices()))
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-offSetDay));
                offSetDay++;
            }
            tabletFlightPage.WaitPageLoading();
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowGuest, true);
            tabletFlightPage.CliqueSurOKFilterBtn();
            string firstflightbefore = tabletFlightPage.GetFirstFlightNumber();
            //2.Cliquer sur...
            DetailFlightTabletAppPage detailFlight = tabletFlightPage.EditFlightHaveServicesExceptXUFlights();
            detailFlight.ShowExtendedMenu();
            detailFlight.ClickBackButton();
            string firstflightafter = tabletFlightPage.GetFirstFlightNumber();
            Assert.AreEqual(firstflightbefore, firstflightafter, "La vue initiale a été changée");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Notification_Template()
        {
            // Prepare
            string DocSite = "ACE";

            // Arrange
            var homePage = LogInAsAdmin();

            FlightTabletAppPage tabletFlightPage = null;

            try
            {
                // Act
                TabletAppPage tabletAppPage = homePage.GotoTabletApp();
                tabletAppPage.Select("mat-select-0", DocSite);
                tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
                tabletFlightPage.ClickSurChooseAViewDefault();
                tabletFlightPage.WaitPageLoading();
                var dataExist = tabletFlightPage.VerifyDataExist();
                if (dataExist == false)
                {
                    tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
                }
                tabletFlightPage.WaitPageLoading();
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, false);
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowNotification, true);
                tabletFlightPage.CliqueSurOKFilterBtn();
                tabletFlightPage.ClickFirstNotification();
                tabletFlightPage.AddNotificationButton();
                tabletFlightPage.AddNotificationText("test");
                tabletFlightPage.AddNotificationAfterText();
                tabletFlightPage.ClickFirstNotification();
                tabletFlightPage.ConfirmNotificationForAdd();
                tabletFlightPage.Savenotificationforadd();
                bool verif = tabletFlightPage.VerifNotif();
                Assert.IsTrue(verif, "Notification ne fonctionne pas");
            }
            finally
            {
                if (tabletFlightPage != null)
                {
                    tabletFlightPage.ClickFirstNotification();
                    tabletFlightPage.DeleteNotif();
                    tabletFlightPage.AcceptJavaScriptConfirmation();
                }
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Favori()
        {
            // Prepare
            string DocSite = "ACE";
            string nameView = "View" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();

            try
            {
                //Delete views added
                tabletFlightPage.ClickSurChooseAView();
                tabletFlightPage.DeleteAllViews();
                tabletFlightPage.UnclickSurChooseAView();

                tabletFlightPage.ClickSurChooseAViewDefault();
                var dataExist = tabletFlightPage.VerifyDataExist();
                if (dataExist == false)
                {
                    tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
                }

                var data = tabletFlightPage.GetFlights();
                string customerName = tabletFlightPage.GetCustomer();

                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowGuest, true);

                //2 Sélectionner Filter -Auto Scroll- Customer
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Customer, "Select All");
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Customer, "UnSelect All");
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Customer, customerName);
                tabletFlightPage.CliqueSurOKFilterBtn();

                //Add new View 
                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.SaveView(nameView);

                var data2 = tabletFlightPage.GetFlights();

                Assert.AreNotEqual(data, data2, "le filtre par favoris doit fonctionner et afficher uniquement les flights liés a ce favoris");
            }
            finally
            {
                tabletFlightPage.ClickSurChooseAView();
                tabletFlightPage.DeleteAllViews();
                tabletFlightPage.UnclickSurChooseAView();

                tabletFlightPage.ClickSurChooseAViewDefault();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_ShowNotification_TimeOut()
        {
            // Prepare
            string DocSite = "ACE";
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            try
            {
                tabletFlightPage.ClickSurChooseAViewDefault();
                var dataExist = tabletFlightPage.VerifyDataExist();
                if (dataExist == false)
                {
                    tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
                }

                tabletFlightPage.CliqueSurFilterIcone();
                tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowNotification, true);
                // Démarrer le chronomètre
                var stopwatch = new Stopwatch();
                stopwatch.Start();
                tabletFlightPage.CliqueSurOKFilterBtn();
                stopwatch.Stop();

                //Assert - Vérification que l'opération ne prend pas trop de temps
                double maxAllowedTime = 6000;
                var checkDelay = stopwatch.ElapsedMilliseconds < maxAllowedTime;
                Assert.IsTrue(checkDelay, $"L'opération a pris trop de temps : {stopwatch.ElapsedMilliseconds}ms, alors que le maximum autorisé est {maxAllowedTime}ms.");
            }
            finally
            {
                tabletFlightPage.ClickSurChooseAViewDefault();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_Workshops_Order()
        {
            // Prepare
            string DocSite = "ACE";
            string history = "History";
            string corte = "Corte";
            string ensambl = "Ensambl";

            //Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Select("mat-select-0", DocSite);

            //Etre sur settings pour cocher juste trois workshops
            var applicationSettings = homePage.GoToApplicationSettings();
            var parametersProduction = applicationSettings.GoToParameters_ProductionPage();
            parametersProduction.WorkshopTab();
            //verify ensamblaje et corte existent
            var exist = parametersProduction.VerifyWorkshopExist("Ensamblaje");
            if (!exist)
            {
                parametersProduction.CreateNewWorkshop(ensambl);
            }
            exist = parametersProduction.VerifyWorkshopExist(corte);
            if (!exist)
            {
                parametersProduction.CreateNewWorkshop(corte);
            }
            exist = parametersProduction.VerifyWorkshopExist(history);
            if (!exist)
            {
                parametersProduction.CreateNewWorkshop(history);
            }

            var parametersTablet = applicationSettings.GoToParametres_Tablet();
            parametersTablet.WorkshopsTab();
            parametersTablet.ShowOnTabletByLabel(ensambl, corte, history);
            parametersTablet.DesactiverTabletByLabel(ensambl, corte, history);

            //Etre sur le module Flight Tablet et avoir des vols
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.ClickSurChooseAViewDefault();
            var dataExist = tabletFlightPage.VerifyDataExist();
            if (dataExist == false)
            {
                tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            }
            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, true);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowFlights_ByflightDate, true);
            //tabletFlightPage.CliqueSurOKFilterBtn();
            tabletFlightPage.clickonTimeBlock();
            tabletFlightPage.WaitPageLoading();

            // test sur order des colonnes Ensambl Corte History 
            bool isVerifiedOrdreWorkshop = tabletFlightPage.VerifiedOrdreWorkshop(ensambl, corte, history);
            //Assert
            Assert.IsTrue(isVerifiedOrdreWorkshop, "L'affichage de l'ordre des WorkShop doit étre le méme paramétré dans Setting Tablet WorkShops (1-2-3,...).");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_ResponsivityFlightView()
        {
            // Prepare
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            string customer = "$$ - CAT Genérico";
            string customer2 = "TVS - SMARTWINGS, A.S.  (TVS)";
            string BCN = TestContext.Properties["SiteToFlight"].ToString();
            string ACE = TestContext.Properties["SiteLP"].ToString();
            string flightNo = "FlightACE_";
            string serviceName = TestContext.Properties["FlightService"].ToString();
            string guestName = TestContext.Properties["Guest"].ToString();
            var rnd = new Random();

            string flight1 = flightNo + "aujourdhui_" + DateUtils.Now.ToString("yyyy-MM-dd") + "_ForTabletApp_" + customer + rnd.Next(); ;
            //Arrange
            var homePage = LogInAsAdmin();

            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            FlightCreateModalPage modal = flightPage.FlightCreatePage();
            modal.FillField_CreatNewFlight(flight1, customer2, aircraft, ACE, BCN, null, "00", "23", null, DateUtils.Now);
            homePage.Navigate();
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            tabletFlightPage.ClickSurChooseAViewDefault();
            //Filter
            tabletFlightPage.WaitLoading();
            tabletFlightPage.WaitLoading();
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowETA, true);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowGuest, true);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowWorkshop, true);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowLeg, true);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowRobotTSU, true);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowDriver, true);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowDock, true);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowNotification, true);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowSpecMeal, true);
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.ShowRegistration, true);
            tabletFlightPage.CliqueSurOKFilterBtn();
            Assert.IsTrue(tabletFlightPage.AreAllColumnTitlesAligned(), "Les colonnes devaient être alignées sur différents formats d'écran, mais ce n'était pas le cas.");
            tabletFlightPage.ScrollToEndOfHorizontalScroll();
            bool resultat = tabletFlightPage.IsHorizontalScrollBarVisible();
            Assert.IsTrue(resultat, "La barre de défilement horizontal doit permettre de voir la totalité des informations.");
        }

        

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_ShowDatasheet()
        {
            //prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            homePage.Navigate();
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", site);

            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Search, flightNumberToday);
            tabletFlightPage.CliqueSurOKFilterBtn();
            tabletFlightPage.WaitLoading();
            tabletFlightPage.EditFirstFlight();

            //vérification du service
            string service = tabletFlightPage.GetFirstServiceName();
            bool serviceExist = serviceName.Contains(service);
            Assert.IsTrue(serviceExist, "Le Service crée ne s'affiche pas");
            tabletFlightPage.ClickFirstDatasheet();

            //vérification du datasheet
            string datasheet = tabletFlightPage.GetFirstDatasheetName();
            Assert.AreEqual(datasheetName, datasheet, "La Datasheet crée ne s'affiche pas");

            //vérification de l'item
            tabletFlightPage.ClickFirstRecipe();
            bool firstItemExist = tabletFlightPage.CheckFirstItem();
            Assert.IsTrue(firstItemExist, "L'item ajouter ne s'affiche pas.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_FLIGHT_StateChangement()
        {
            //prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            homePage.Navigate();
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", site);

            FlightTabletAppPage tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Search, flightNumberToday);
            tabletFlightPage.CliqueSurOKFilterBtn();
            tabletFlightPage.WaitLoading();
            bool isFlightValidated = tabletFlightPage.IsFlightValidatedOnly();

            //Assert
            Assert.IsFalse(isFlightValidated, "Le changement de statut P n' est pas bien pris en compte sous Tablet");

            //Validate flight
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberToday);
            flightPage.SetNewState("V");

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", site);
            tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Search, flightNumberToday);
            tabletFlightPage.CliqueSurOKFilterBtn();
            tabletFlightPage.WaitLoading();
            isFlightValidated = tabletFlightPage.IsFlightValidatedOnly();

            //Assert
            Assert.IsTrue(isFlightValidated, "Le changement de statut V n' est pas bien pris en compte sous Tablet");

            //Invoiced flight
            homePage.Navigate();
            flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberToday);
            flightPage.SetNewState("I");

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", site);
            tabletFlightPage = tabletAppPage.GotoTabletApp_Flight();
            tabletFlightPage.SetDate(DateTime.UtcNow.AddDays(-1));
            //Filter
            tabletFlightPage.CliqueSurFilterIcone();
            tabletFlightPage.Filter(FlightTabletAppPage.FilterType.Search, flightNumberToday);
            tabletFlightPage.CliqueSurOKFilterBtn();
            tabletFlightPage.WaitLoading();
            bool isFlightInvoiced = tabletFlightPage.IsInvoiced();

            //Assert
            Assert.IsTrue(isFlightInvoiced, "Le changement de statut I n' est pas bien pris en compte sous Tablet");

        }

        ////////////////////////////////////// Insert ///////////////////////////////
        private int InsertFlightDeliveries(string site, string customer, string flightName)
        {
            //Insert FlightDeliveries For Flight
            string queryFlightDeliveries = @"
            DECLARE @siteId INT;
            SELECT TOP 1 @siteId = Id FROM sites WHERE Name LIKE @site;

            DECLARE @customerId INT;
            SELECT TOP 1 @customerId = Id FROM customers WHERE Name LIKE @customer;

            Insert into [FlightDeliveries](
              [Name]
              ,[CustomerId]
              ,[SiteId]
              ,[IsActive]
              ,[DeliveryTime]
              ,[HoursBeforeAccess]
              ,[AllowedModificationsPercent]
              ,[HoursBeforeModifications]
              ,[Method]
              ,[CustomerPortalBlockAccessType])
	          values (
	          @flightName,
	          @customerId,
	          @siteId,
	          1,
	          '00:00:00.0000000',
	          '5',
	          '0',
	          '0',
	          '45',
	          '10'
	          );
 SELECT SCOPE_IDENTITY();
"
            ;
            return ExecuteAndGetInt(
            queryFlightDeliveries,
                new KeyValuePair<string, object>("site", site),
                new KeyValuePair<string, object>("customer", customer),
                new KeyValuePair<string, object>("flightName", flightName)
            );
        }

        private void InsertFlight(int flightDeliverieId, string date, string aircraft, string airoportFrom, string airoportTo)
        {
            //Insert Flight
            string queryFlight = @"
            DECLARE @aircraftId INT;
            SELECT TOP 1 @aircraftId = Id FROM Aircraft WHERE Name LIKE @aircraft;

            DECLARE @airoportFromId INT;
            SELECT TOP 1 @airoportFromId = Id FROM Airports WHERE Name LIKE @airoportFrom;

            DECLARE @airoportToId INT;
            SELECT TOP 1 @airoportToId = Id FROM Airports WHERE Name LIKE @airoportTo;

            INSERT INTO flight (
            [Id],
            [AirportFromId],
            [AirportToId],
            [ETA],
            [ETD],
            [FlightType],
            [IsPreval],
            [IsValid],
            [IsToInvoice],
            [IsCancelled],
            [FlightDate],
            [AircraftId],
            [IsLocked],
            [ProductionStatus],
            [FinalState],
            [IsUpdatedByTP006],
            [IsMajor],
            [NbServToProduceIsWarning],
            [TSURobotProdStatus]
        ) VALUES (
            @flightDeliverieId,
            @airoportFromId,
            @airoportToId,
            '00:00:00.0000000',
            '23:00:00.0000000',
            '1',
            1,
            0,
            0,
            0,
            @date,
            @aircraftId,
            0,
            '24',
            '45',
            0,
            0,
	        0,
            '45'
        );
        "
            ;
            ExecuteNonQuery(
            queryFlight,
                new KeyValuePair<string, object>("flightDeliverieId", flightDeliverieId),
                new KeyValuePair<string, object>("date", date),
                new KeyValuePair<string, object>("aircraft", aircraft),
                new KeyValuePair<string, object>("airoportFrom", airoportFrom),
                new KeyValuePair<string, object>("airoportTo", airoportTo)
            );
        }

        private int InsertFlightLegs(int flightId)
        {
            //Insert FlightLegs For Flight
            string queryFlightLegs = @"Insert into [FlightLegs] 
                ([FlightId]
              ,[AirportFromId]
              ,[AirportToId]
              ,[Order])
	          values
	          (
	          @flightId,
	          '11',
	          '10',
	          '254'
	          );
 SELECT SCOPE_IDENTITY();
                "
            ;
            return ExecuteAndGetInt(
            queryFlightLegs,
                new KeyValuePair<string, object>("flightId", flightId)
            );
        }

        private int InsertFlightLegToGuestTypes(string guestName, int flightLegId)
        {
            //Insert FlightLegToGuestTypes For Flight
            string queryFlightLegToGuestTypes = @"
            DECLARE @guestId INT;
            SELECT TOP 1 @guestId = Id FROM GuestTypes WHERE Name LIKE @guestName;

               Insert into [FlightLegToGuestTypes] 
               ([FlightLegId]
              ,[GuestTypeId]
              ,[NbGuestTypes]
              ,[MenuCodeManuallyUpdated])
	          values
	          (
	          @flightLegId,
	          @guestId,
	          '2',
	          0
	          );
 SELECT SCOPE_IDENTITY();
                "
            ;
            return ExecuteAndGetInt(
            queryFlightLegToGuestTypes,
                new KeyValuePair<string, object>("guestName", guestName),
                new KeyValuePair<string, object>("flightLegId", flightLegId)
            );
        }

        private int InsertFlightDetailsWithService(string serviceName, int flightLegToGuestTypeId)
        {
            if (string.IsNullOrWhiteSpace(serviceName))
            {
                throw new ArgumentException("Le nom du service (serviceName) ne peut pas être null ou vide.");
            }

            string queryFlightDetails = @"
        DECLARE @serviceId INT;

        -- Récupération du ServiceId
        SELECT TOP 1 @serviceId = Id 
        FROM Services 
        WHERE Name LIKE @serviceName;

        -- Si aucun service n'est trouvé, lever une erreur explicite
        IF @serviceId IS NULL
        BEGIN
            THROW 51000, 'ServiceId introuvable pour le service spécifié.', 1;
        END

        -- Insertion dans FlightDetails
        INSERT INTO [FlightDetails] (
            [FlightLegToGuestTypeId],
            [ServiceId],
            [Mode],
            [Value],
            [FinalQuantity],
            [ComputedQuantity],
            [CreationDate],
            [ModificationDate],
            [ProductionStatus],
            [IsManual],
            [IsSpml]
        )
        VALUES (
            @flightLegToGuestTypeId,
            @serviceId,
            '10',
            '10',
            '10',
            '10',
            '2016-06-18 00:00:00.000',
            '2016-06-18 00:00:00.000',
            '10',
            0,
            0
        );

        -- Retourner l'identifiant généré
        SELECT SCOPE_IDENTITY();
    ";

            return ExecuteAndGetInt(
                queryFlightDetails,
                new KeyValuePair<string, object>("serviceName", serviceName),
                new KeyValuePair<string, object>("flightLegToGuestTypeId", flightLegToGuestTypeId)
            );
        }

        /// <summary>
        /// ///////////////////////// Delete /////////////////////////////
        /// </summary>
        private void DeleteFlightDetailsWithService(int flightDetailsId)
        {
            string query = @"
            DELETE FROM FlightDetails 
                    WHERE Id = @flightDetailsId;";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("flightDetailsId", flightDetailsId));
        }

        private void DeleteFlightLegToGuestType(int flightLegToGuestTypesId)
        {
            string query = @"
            DELETE FROM FlightLegToGuestTypes 
                    WHERE Id = @flightLegToGuestTypesId;";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("flightLegToGuestTypesId", flightLegToGuestTypesId));
        }

        private void DeleteFlightLeg(int flightLegsId)
        {
            string query = @"
        DELETE FROM FlightLegs 
        WHERE Id = @flightLegsId;";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("flightLegsId", flightLegsId));
        }

        private void DeleteFlight(int flightId)
        {
            string query = @"
        DELETE FROM flight 
        WHERE Id = @flightId;";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("flightId", flightId));
        }

        private void DeleteFlightPropertie(int flightId)
        {
            string query = @"
        DELETE FROM FlightProperties 
        WHERE FlightId = @flightId;";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("flightId", flightId));
        }

        private void DeleteFlightNotificationStatusChange(int flightId)
        {
            string query = @"
            delete from FlightNotificationStatusChanges where FlightId = @flightId";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("flightId", flightId));
        }

        private void DeleteFlightDeliverie(int flightdeliveryId)
        {
            string query = @"
        DELETE FROM FlightDeliveries 
        WHERE Id = @flightdeliveryId;";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("flightdeliveryId", flightdeliveryId));
        }

        private void SetTabletFlightType(string type)
        {
            string query = @"
            DECLARE @tabletFlightTypeId INT;
            SELECT TOP 1 @tabletFlightTypeId = Id FROM Services WHERE Name LIKE @type;

            UPDATE [dbo].[Flight] 
            SET TabletFlightTypeId = @tabletFlightTypeId 
            WHERE id = @tabletFlightTypeId;";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("type", type));
        }

        private void DeleteFlightHistorie(int flightId)
        {
            string query = @"
            delete from FlightHistories where FlightId = @flightId;";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("flightId", flightId));
        }
    }
}