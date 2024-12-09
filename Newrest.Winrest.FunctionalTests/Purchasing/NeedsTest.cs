using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Schedule;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Needs;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Scheduler;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Policy;
using System.Web;

namespace Newrest.Winrest.FunctionalTests.Purchasing
{
    [TestClass]
    public class NeedsTests : TestBase
    {

        private const int _timeout = 600000;
        private readonly string RAW_MAT_GROUP = "RawMatByGroup";
        private readonly string RAW_MAT_WORKSHOP = "RawMatByWorkshop";
        private readonly string RAW_MAT_SUPPLIER = "RawMatBySupplier";
        private readonly string RAW_MAT_RECIPE = "RawMatByRecipe";
        private readonly string RAW_MAT_CUSTOMER = "RawMatByCustomer";
        private readonly string GROUP_BY_ITEMGROUP = "ItemGroup";
        private readonly string GROUP_BY_WORKSHOP = "Workshop";
        private readonly string GROUP_BY_RECIPE_TYPE = "RecipeType";
        private readonly string GROUP_BY_CUSTOMER = "Customer";

        #region "Test init & Test Clean Up"
        [TestInitialize]
        public override void TestInitialize()
        {

            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(PU_NEEDS_RawMatByCustomer):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PU_NEEDS_Filter_Customers):
                    Test_Initialize_PU_NEEDS_CreateFlight_Validated_flight_CUSTOMER3_Valid_J1();
                    break;
                case nameof(PU_NEEDS_Filter_RecipeType):
                    Test_Initialize_PU_NEEDS_CreateFlight_Validated_flight_CUSTOMER3_Valid_J1();
                    break;
                case nameof(PU_NEEDS_Filter_Service):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    break;

                case nameof(PU_NEEDS_Filter_ServiceCategorie):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PU_NEEDS_Filter_ShowNormalAndVacuumProduction):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    break;
                case nameof(PU_NEEDS_Filter_ShowNormalProductionOnly):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    break;
                case nameof(PU_NEEDS_Filter_ShowVacuumProductionOnly):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    break;
                case nameof(PU_NEEDS_Filter_Workshops):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    Test_Initialize_PU_NEEDS_CreateFlight_Validated_flight_CUSTOMER3_Valid_J1();
                    break;
                case nameof(PU_NEEDS_UnfoldAll):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PU_NEEDS_Filter_Customers_FaF):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    break;

                case nameof(PU_NEEDS_GenerateSupplyOrder):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PU_NEEDS_Filter_RecipeType_FaF):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PU_NEEDS_Filter_Site):
                    Test_Initialize_PR_PRODMAN_CreateFlight_MAD_ServiceNameNormalMAD();
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PU_NEEDS_GroupBy_RawMatByGroup):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PU_NEEDS_RawMatBySupplier):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    Test_Initialize_PU_NEEDS_CreateFlight_Validated_flight_CUSTOMER3_GUEST2_Valid_J1();
                    break;
                case nameof(PU_NEEDS_ResetAllQuantities):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    Test_Initialize_PU_NEEDS_CreateFlight_Validated_flight_CUSTOMER3_GUEST2_Valid_J1();
                    break;
                case nameof(PU_NEEDS_Filter_GuestType):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    Test_Initialize_PU_NEEDS_CreateFlight_Validated_flight_CUSTOMER3_GUEST2_Valid_J1();
                    break;
                case nameof(PU_NEEDS_Filter_ShowValidatedFlightsOnly):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    Test_Initialize_PU_NEEDS_CreateFlight_Validated_flight_CUSTOMER3_GUEST2_Valid_J1();
                    break;
                case nameof(PU_NEEDS_Filter_Service_FaF):
                    Test_Initialize_PR_PRODMAN_CreateFlight_MAD_ServiceNameNormalMAD();
                    break;
                case nameof(PU_NEEDS_Filter_ItemGroup_FaF):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PU_NEEDS_Filter_SortByCustomer):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    break;
                case nameof(PU_NEEDS_Filter_SortByService):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    break;
                case nameof(PU_NEEDS_Filter_ServiceCategorie_FaF):
                    Test_Initialize_PR_PRODMAN_CreateFlight_MAD_ServiceNameNormalMAD();
                    break;
                case nameof(PU_NEEDS_Filter_ShowNormalAndVacuumProduction_FaF):
                    Test_Initialize_PR_PRODMAN_CreateFlight_MAD_ServiceNameNormalMAD();
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    break;
                case nameof(PU_NEEDS_FoldAll):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PU_NEEDS_Filter_From_Result):
                    Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ0();
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
                case nameof(PU_NEEDS_RawMatByCustomer):
                case nameof(PU_NEEDS_Filter_Customers):
                case nameof(PU_NEEDS_Filter_RecipeType):
                case nameof(PU_NEEDS_Filter_Service):
                case nameof(PU_NEEDS_Filter_ServiceCategorie):
                case nameof(PU_NEEDS_Filter_ShowNormalAndVacuumProduction):
                case nameof(PU_NEEDS_Filter_ShowNormalProductionOnly):
                case nameof(PU_NEEDS_Filter_ShowVacuumProductionOnly):
                case nameof(PU_NEEDS_Filter_Workshops):
                case nameof(PU_NEEDS_UnfoldAll):
                case nameof(PU_NEEDS_Filter_Customers_FaF):
                case nameof(PU_NEEDS_GenerateSupplyOrder):
                case nameof(PU_NEEDS_Filter_RecipeType_FaF):
                case nameof(PU_NEEDS_Filter_Site):
                case nameof(PU_NEEDS_GroupBy_RawMatByGroup):
                case nameof(PU_NEEDS_RawMatBySupplier):
                case nameof(PU_NEEDS_ResetAllQuantities):
                case nameof(PU_NEEDS_Filter_GuestType):
                case nameof(PU_NEEDS_Filter_ShowValidatedFlightsOnly):
                case nameof(PU_NEEDS_Filter_Service_FaF):
                case nameof(PU_NEEDS_Filter_ItemGroup_FaF):
                case nameof(PU_NEEDS_Filter_SortByCustomer):
                case nameof(PU_NEEDS_Filter_SortByService):
                case nameof(PU_NEEDS_Filter_ServiceCategorie_FaF):
                case nameof(PU_NEEDS_Filter_ShowNormalAndVacuumProduction_FaF):
                case nameof(PU_NEEDS_FoldAll):
                case nameof(PU_NEEDS_Filter_From_Result):
                    TB_FLIGHT_TestCleanUp();
                    break;
                default:
                    break;
            }
            base.TestCleanup();
        }
        private void TB_FLIGHT_TestCleanUp()
        {
            //// delete FlightDetails
            List<int> FlightDetails = new List<int>();

            for (int i = 1; i <= 12; i++)
            {
                if (TestContext.Properties.Contains($"FlightDetailsId{i}"))
                {
                    FlightDetails.Add((int)TestContext.Properties[$"FlightDetailsId{i}"]);
                }
            }
            DeleteFlightDetails(FlightDetails);

            //// delete FlightLegToGuestTypes
            List<int> FlightLegToGuestTypes = new List<int>();

            for (int i = 1; i <= 12; i++)
            {
                if (TestContext.Properties.Contains($"FlightFlightLegToGuestTypesId{i}"))
                {
                    FlightLegToGuestTypes.Add((int)TestContext.Properties[$"FlightFlightLegToGuestTypesId{i}"]);
                }
            }
            DeleteFlightLegToGuestTypes(FlightLegToGuestTypes);

            //// delete FlightLegs
            List<int> FlightLegs = new List<int>();

            for (int i = 1; i <= 12; i++)
            {
                if (TestContext.Properties.Contains($"FlightLegsId{i}"))
                {
                    FlightLegs.Add((int)TestContext.Properties[$"FlightLegsId{i}"]);
                }
            }
            DeleteFlightLegs(FlightLegs);


            #region deleteFlightProperties
            List<int> FlightForProperties = new List<int>();

            for (int i = 1; i <= 12; i++)
            {
                if (TestContext.Properties.Contains($"FlightDeliveriesId{i}"))
                {
                    FlightForProperties.Add((int)TestContext.Properties[$"FlightDeliveriesId{i}"]);
                }
            }
            DeleteFlightProperties(FlightForProperties);
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


            //// delete FlightDeliveries
            List<int> Flights = new List<int>();

            for (int i = 1; i <= 12; i++)
            {
                if (TestContext.Properties.Contains($"FlightDeliveriesId{i}"))
                {
                    Flights.Add((int)TestContext.Properties[$"FlightDeliveriesId{i}"]);
                }
            }
            DeleteFlights(Flights);

            //// delete FlightDeliveries
            List<int> FlightDeliveries = new List<int>();

            for (int i = 1; i <= 12; i++)
            {
                if (TestContext.Properties.Contains($"FlightDeliveriesId{i}"))
                {
                    FlightDeliveries.Add((int)TestContext.Properties[$"FlightDeliveriesId{i}"]);
                }
            }
            DeleteFlightDeliveries(FlightDeliveries);
        }
        #endregion
        #region "Test Initi creat flight "
        public void Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ1()
        {

            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string flightNameNormal = "" + TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.AddDays(1).ToString("dd/MM/yyyy");

            TestContext.Properties[string.Format("FlightDeliveriesId1")] = InsertFlightDeliveries(siteACE, customerName1, flightNameNormal);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId1")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId1")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId1")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId1")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId1")]);
            TestContext.Properties[string.Format("FlightDetailsId1")] = InsertFlightDetailsWithService(serviceNameNormal, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId1")]);

        }
        public void Test_Initialize_PU_NEEDS_CreateFlight_normal_flightJ0()
        {

            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string flightNameNormal = "" + TestContext.Properties["Needs_FlightNormal"].ToString() + DateTime.Today.AddDays(-1).ToString("dd/MM/yyyy");

            TestContext.Properties[string.Format("FlightDeliveriesId1")] = InsertFlightDeliveries(siteACE, customerName1, flightNameNormal);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId1")], DateTime.Today.AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId1")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId1")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId1")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId1")]);
            TestContext.Properties[string.Format("FlightDetailsId1")] = InsertFlightDetailsWithService(serviceNameNormal, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId1")]);

        }
        public void Test_Initialize_PU_NEEDS_CreateFlight_Validated_flight_CUSTOMER3_Valid_J1()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();

            string customerName3 = TestContext.Properties["Needs_Customer3"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();
            string flightNameValide = "" + TestContext.Properties["Needs_FlightValide"].ToString() + DateUtils.Now.AddDays(1).ToString("dd/MM/yyyy");


            TestContext.Properties[string.Format("FlightDeliveriesId2")] = InsertFlightDeliveries(siteACE, customerName3, flightNameValide);
            InsertValidFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId2")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId2")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId2")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId2")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId2")]);
            TestContext.Properties[string.Format("FlightDetailsId2")] = InsertFlightDetailsWithService(serviceNameFlightValide, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId2")]);

        }
        public void Test_Initialize_PU_NEEDS_CreateFlight_Validated_flight_CUSTOMER3_GUEST2_Valid_J1()
        {

            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();


            string customerName3 = TestContext.Properties["Needs_Customer3"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();
            string flightNameValide = "" + TestContext.Properties["Needs_FlightValide"].ToString() + DateUtils.Now.AddDays(1).ToString("dd/MM/yyyy");


            TestContext.Properties[string.Format("FlightDeliveriesId3")] = InsertFlightDeliveries(siteACE, customerName3, flightNameValide);
            InsertValidFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId3")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId3")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId3")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId3")] = InsertFlightLegToGuestTypes(guestType2, (int)TestContext.Properties[string.Format("FlightLegsId3")]);
            TestContext.Properties[string.Format("FlightDetailsId3")] = InsertFlightDetailsWithService(serviceNameFlightValide, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId3")]);

        }
        public void Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();


            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName2 = TestContext.Properties["Needs_Customer2"].ToString();
            string serviceNameVacuum = TestContext.Properties["Needs_ServiceVacuum"].ToString();
            string flightNameVacuum = "" + TestContext.Properties["Needs_FlightVacuum"].ToString() + DateUtils.Now.AddDays(1).ToString("dd/MM/yyyy");
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();

            TestContext.Properties[string.Format("FlightDeliveriesId4")] = InsertFlightDeliveries(siteACE, customerName2, flightNameVacuum);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId4")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId4")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId4")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId4")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId4")]);
            TestContext.Properties[string.Format("FlightDetailsId4")] = InsertFlightDetailsWithService(serviceNameVacuum, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId4")]);

        }
        public void Test_Initialize_PR_PRODMAN_CreateFlight_MAD_ServiceNameNormalMAD()
        {

            string siteMAD = TestContext.Properties["Site"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormalMAD = TestContext.Properties["Needs_ServiceNormalMAD"].ToString();
            string flightNameNormalMAD = "" + TestContext.Properties["Needs_FlightNormalMAD"].ToString() + DateUtils.Now.AddDays(1).ToString("dd/MM/yyyy");


            TestContext.Properties[string.Format("FlightDeliveriesId5")] = InsertFlightDeliveries(siteMAD, customerName1, flightNameNormalMAD);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId5")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteMAD, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId5")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId5")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId5")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId5")]);
            TestContext.Properties[string.Format("FlightDetailsId5")] = InsertFlightDetailsWithService(serviceNameNormalMAD, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId5")]);

        }
        #endregion

      
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void PU_NEEDS_SetConfigWinrest4_0()
        {
            //Prepare
            var keyword = TestContext.Properties["Item_Keyword"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // New version search
            homePage.SetNewVersionSearchValue(true);

            // New keyword search
            homePage.SetNewVersionKeywordValue(true);

            // New VersionItemDetails
            //homePage.SetVersionItemDetailValue(true);

            // New group display
            homePage.SetNewGroupDisplayValue(true);

            var itemPage = homePage.GoToPurchasing_ItemPage();

            // Vérifier New version search
            try
            {
                var firstItem = itemPage.GetFirstItemName();
                itemPage.Filter(ItemPage.FilterType.Search, firstItem);
            }
            catch
            {
                throw new Exception("La recherche n'a pas pu être effectuée, le NewSearchMode est inactif.");
            }

            // vérifier new keyword search
            try
            {
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Keyword, keyword);
            }
            catch
            {
                throw new Exception("La recherche par keyword n'a pas pu être effectuée, le NewKeywordMode est inactif.");
            }

            // vérifier new version item detail
            try
            {
                itemPage.ResetFilter();
                var itemDetails = itemPage.ClickOnFirstItem();
                itemDetails.SearchPackaging(site);
            }
            catch
            {
                throw new Exception("La recherche d'un packaging n'a pas pu être effectuée, le NewVersionItemDetails est inactif.");
            }

            // vérifier new group display
            ReceiptNotesItem receiptNotesItem;
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            if (receiptNotesPage.CheckTotalNumber() == 0)
            {
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, delivery));
                receiptNotesItem = receiptNotesCreateModalpage.Submit();
            }
            else
            {
                receiptNotesItem = receiptNotesPage.SelectFirstReceiptNoteItem();
            }

            Assert.IsTrue(receiptNotesItem.IsGroupDisplayActive(), "Le paramètre 'NewGroupDisplay' n'est pas activé.");
        }
        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_NEEDS_CreateItems()
        {
            //Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string itemName1 = TestContext.Properties["Needs_Item1"].ToString();
            string itemName2 = TestContext.Properties["Needs_Item2"].ToString();
            string itemName3 = TestContext.Properties["Needs_Item3"].ToString();
            string itemName4 = TestContext.Properties["Needs_Item4"].ToString();
            string group1 = TestContext.Properties["Prodman_Needs_Item_Group1"].ToString();
            string group2 = TestContext.Properties["Prodman_Needs_Item_Group2"].ToString();
            string group3 = TestContext.Properties["Prodman_Needs_Item_Group3"].ToString();
            string group4 = TestContext.Properties["Prodman_Needs_Item_Group4"].ToString();
            string workshop1 = TestContext.Properties["Prodman_Needs_Workshop1"].ToString();
            string workshop2 = TestContext.Properties["Prodman_Needs_Workshop2"].ToString();
            string workshop3 = TestContext.Properties["Prodman_Needs_Workshop3"].ToString();
            string supplier1 = TestContext.Properties["Prodman_Needs_Item_Supplier1"].ToString();
            string supplier2 = TestContext.Properties["Prodman_Needs_Item_Supplier2"].ToString();
            string supplier3 = TestContext.Properties["Prodman_Needs_Item_Supplier3"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();

            // ----------------------------------------- Item 1 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName1, group1, workshop1, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier1);

                // 1 packaging pour le site MAD
                itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteMAD, packagingName, storageQty, storageUnit, qty, supplier1);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName1);
            }

            string item1 = itemPage.GetFirstItemName();

            // ----------------------------------------- Item 2 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName2);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName2, group2, workshop2, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier2);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName2);
            }

            string item2 = itemPage.GetFirstItemName();

            // ----------------------------------------- Item 3 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName3);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName3, group3, workshop3, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier3);

                // 1 packaging pour le site MAD
                itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteMAD, packagingName, storageQty, storageUnit, qty, supplier3);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName3);
            }

            string item3 = itemPage.GetFirstItemName();

            // ----------------------------------------- Item 4 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName4);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName4, group4, workshop1, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier1);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName4);
            }

            string item4 = itemPage.GetFirstItemName();

            // Assert
            Assert.AreEqual(itemName1, item1, "L'item " + itemName1 + " n'existe pas.");
            Assert.AreEqual(itemName2, item2, "L'item " + itemName2 + " n'existe pas.");
            Assert.AreEqual(itemName3, item3, "L'item " + itemName3 + " n'existe pas.");
            Assert.AreEqual(itemName4, item4, "L'item " + itemName4 + " n'existe pas.");
        }

        [Priority(2)]
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_NEEDS_CreateRecipe()
        {
            //Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string recipeVariant = TestContext.Properties["Prodman_Needs_RecipeVariant"].ToString();
            string recipeName1 = TestContext.Properties["Needs_RecipeName1"].ToString();
            string recipeType1 = TestContext.Properties["Prodman_Needs_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Prodman_Needs_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Prodman_Needs_Workshop1"].ToString();
            string recipeName2 = TestContext.Properties["Needs_RecipeName2"].ToString();
            string recipeType2 = TestContext.Properties["Prodman_Needs_RecipeType2"].ToString();
            string recipeCookingMode2 = TestContext.Properties["Prodman_Needs_CookingMode2"].ToString();
            string recipeWorkshop2 = TestContext.Properties["Prodman_Needs_Workshop2"].ToString();
            string recipeName3 = TestContext.Properties["Needs_RecipeName3"].ToString();
            string recipeType3 = TestContext.Properties["Prodman_Needs_RecipeType3"].ToString();
            string recipeWorkshop3 = TestContext.Properties["Prodman_Needs_Workshop3"].ToString();
            string itemName1 = TestContext.Properties["Needs_Item1"].ToString();
            string itemName2 = TestContext.Properties["Needs_Item2"].ToString();
            string itemName3 = TestContext.Properties["Needs_Item3"].ToString();
            string itemName4 = TestContext.Properties["Needs_Item4"].ToString();
            Random rnd = new Random();

            //Arrange           
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();

            // ---------------------------------------------- recette 1 --------------------------------------------------
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName1);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName1, recipeType1, rnd.Next(1, 30).ToString(), true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE et MAD
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariant);
                recipeGeneralInfosPage.AddVariantWithSite(siteMAD, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName1);
                recipeVariantPage.AddIngredient(itemName3);
                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemName1);
                    recipeVariantPage.AddIngredient(itemName3);

                }
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName1);

            string recipe1 = recipesPage.GetFirstRecipeName();

            // ---------------------------------------------- recette 2 --------------------------------------------------

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName2);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName2, recipeType2, rnd.Next(1, 30).ToString(), true, null, recipeCookingMode2, recipeWorkshop2);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName2);
                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemName2);

                }
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName2);
            string recipe2 = recipesPage.GetFirstRecipeName();

            // ---------------------------------------------- recette 3 --------------------------------------------------

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName3);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName3, recipeType3, rnd.Next(1, 30).ToString(), true, null, recipeCookingMode1, recipeWorkshop3);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName4);
                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemName4);

                }
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName3);
            string recipe3 = recipesPage.GetFirstRecipeName();

            // --------------------------------------------- Assert ------------------------------------------------------------
            Assert.AreEqual(recipeName1, recipe1, "La recette " + recipeName1 + " n'existe pas.");
            Assert.AreEqual(recipeName2, recipe2, "La recette " + recipeName2 + " n'existe pas.");
            Assert.AreEqual(recipeName3, recipe3, "La recette " + recipeName3 + " n'existe pas.");
        }

        [Priority(3)]
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_NEEDS_CreateDatasheet()
        {
            //Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string datasheetGuestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName1 = TestContext.Properties["Needs_RecipeName1"].ToString();
            string recipeName2 = TestContext.Properties["Needs_RecipeName2"].ToString();
            string recipeName3 = TestContext.Properties["Needs_RecipeName3"].ToString();
            string datasheetName1 = TestContext.Properties["Needs_DatasheetName1"].ToString();
            string datasheetNameMAD = TestContext.Properties["Needs_DatasheetNameMAD"].ToString();
            string datasheetName2 = TestContext.Properties["Needs_DatasheetName2"].ToString();
            string datasheetName3 = TestContext.Properties["Needs_DatasheetName3"].ToString();

            //Arrange

            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();

            // ---------------------------------------------- datasheet 1 --------------------------------------------------
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName1);

            if (datasheetPage.CheckTotalNumber() == 0)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName1, datasheetGuestType, siteACE);
                datasheetDetailPage.AddRecipe(recipeName1);
                datasheetDetailPage.BackToList();
            }
            else
            {
                var datasheetDetailPage = datasheetPage.SelectFirstDatasheet();

                if (!datasheetDetailPage.IsRecipeAdded())
                {
                    datasheetDetailPage.AddRecipe(recipeName1);
                }

                datasheetDetailPage.BackToList();
            }

            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName1);
            string datasheet1 = datasheetPage.GetFirstDatasheetName();

            // ---------------------------------------------- datasheet MAD --------------------------------------------------

            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetNameMAD);

            if (datasheetPage.CheckTotalNumber() == 0)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetNameMAD, datasheetGuestType, siteMAD);
                datasheetDetailPage.AddRecipe(recipeName1);
                datasheetDetailPage.BackToList();
            }
            else
            {
                var datasheetDetailPage = datasheetPage.SelectFirstDatasheet();

                if (!datasheetDetailPage.IsRecipeAdded())
                {
                    datasheetDetailPage.AddRecipe(recipeName1);
                }

                datasheetDetailPage.BackToList();
            }

            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetNameMAD);

            string datasheetMAD = datasheetPage.GetFirstDatasheetName();

            // ---------------------------------------------- datasheet 2 --------------------------------------------------

            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName2);

            if (datasheetPage.CheckTotalNumber() == 0)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName2, datasheetGuestType, siteACE);
                datasheetDetailPage.AddRecipe(recipeName2);
                datasheetDetailPage.BackToList();
            }
            else
            {
                var datasheetDetailPage = datasheetPage.SelectFirstDatasheet();

                if (!datasheetDetailPage.IsRecipeAdded())
                {
                    datasheetDetailPage.AddRecipe(recipeName2);
                }

                datasheetDetailPage.BackToList();
            }

            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName2);

            string datasheet2 = datasheetPage.GetFirstDatasheetName();

            // ---------------------------------------------- datasheet 3 --------------------------------------------------

            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName3);

            if (datasheetPage.CheckTotalNumber() == 0)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName3, datasheetGuestType, siteACE);
                datasheetDetailPage.AddRecipe(recipeName3);
                datasheetDetailPage.BackToList();
            }
            else
            {
                var datasheetDetailPage = datasheetPage.SelectFirstDatasheet();

                if (!datasheetDetailPage.IsRecipeAdded())
                {
                    datasheetDetailPage.AddRecipe(recipeName3);
                }

                datasheetDetailPage.BackToList();
            }

            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName3);

            string datasheet3 = datasheetPage.GetFirstDatasheetName();

            // -------------------------------------------------- Assert -------------------------------------------------
            Assert.AreEqual(datasheetName1, datasheet1, "La datasheet " + datasheetName1 + " n'existe pas.");
            Assert.AreEqual(datasheetNameMAD, datasheetMAD, "La datasheet " + datasheetNameMAD + " n'existe pas.");
            Assert.AreEqual(datasheetName2, datasheet2, "La datasheet " + datasheetName2 + " n'existe pas.");
            Assert.AreEqual(datasheetName3, datasheet3, "La datasheet " + datasheetName3 + " n'existe pas.");
        }
        [Priority(4)]
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_NEEDS_CreateCustomer()
        {
            // Prepare
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string customerCode1 = TestContext.Properties["Needs_CustomerCode1"].ToString();
            string customerName2 = TestContext.Properties["Needs_Customer2"].ToString();
            string customerCode2 = TestContext.Properties["Needs_CustomerCode2"].ToString();
            string customerName3 = TestContext.Properties["Needs_Customer3"].ToString();
            string customerCode3 = TestContext.Properties["Needs_CustomerCode3"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();


            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();

            // -------------------------------------------- Customer 1 --------------------------------------------------
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerName1);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName1, customerCode1, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.Filter(CustomerPage.FilterType.Search, customerName1);
            }

            string customer1 = customersPage.GetFirstCustomerName();

            // -------------------------------------------- Customer 2 --------------------------------------------------
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerName2);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName2, customerCode2, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.Filter(CustomerPage.FilterType.Search, customerName2);
            }

            string customer2 = customersPage.GetFirstCustomerName();

            // -------------------------------------------- Customer 3 --------------------------------------------------
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerName3);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName3, customerCode3, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.Filter(CustomerPage.FilterType.Search, customerName3);
            }

            string customer3 = customersPage.GetFirstCustomerName();

            // -------------------------------------------------- Assert -------------------------------------------------
            Assert.AreEqual(customerName1, customer1, "Le customer " + customerName1 + " n'existe pas.");
            Assert.AreEqual(customerName2, customer2, "Le customer " + customerName2 + " n'existe pas.");
            Assert.AreEqual(customerName3, customer3, "Le customer " + customerName3 + " n'existe pas.");
        }

        [TestMethod]
        [Priority(5)]
        [Timeout(_timeout)]

        public void PU_NEEDS_CreateService()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string customerName2 = TestContext.Properties["Needs_Customer2"].ToString();
            string customerName3 = TestContext.Properties["Needs_Customer3"].ToString();
            string datasheetName1 = TestContext.Properties["Needs_DatasheetName1"].ToString();
            string datasheetNameMAD = TestContext.Properties["Needs_DatasheetNameMAD"].ToString();
            string datasheetName2 = TestContext.Properties["Needs_DatasheetName2"].ToString();
            string datasheetName3 = TestContext.Properties["Needs_DatasheetName3"].ToString();
            string serviceCategorie1 = TestContext.Properties["Prodman_Needs_ServiceCategory1"].ToString();
            string serviceCategorie2 = TestContext.Properties["Prodman_Needs_ServiceCategory2"].ToString();
            string serviceCategorie3 = TestContext.Properties["Prodman_Needs_ServiceCategory3"].ToString();
            string serviceType = TestContext.Properties["Prodman_Needs_ServiceType"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameNormalMAD = TestContext.Properties["Needs_ServiceNormalMAD"].ToString();
            string serviceNameVacuum = TestContext.Properties["Needs_ServiceVacuum"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();
            string serviceNameSansDatasheet = TestContext.Properties["Needs_ServiceSansDatasheet"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            // -------------------------------------------- service normal --------------------------------------------------
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameNormal);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceNameNormal, null, null, serviceCategorie1, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerName1, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2), "20", datasheetName1);
                servicePage = pricePage.BackToList();
            }
            else
            {
                var servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.SearchPriceForCustomer(customerName1, siteACE, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2), "20", datasheetName1);

                var serviceGeneralInformationsPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameNormal);

            string serviceNormal = servicePage.GetFirstServiceName();

            // -------------------------------------------- service normal MAD --------------------------------------------------

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameNormalMAD);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceNameNormalMAD, null, null, serviceCategorie1, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteMAD, customerName1, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2), "20", datasheetNameMAD);
                servicePage = pricePage.BackToList();
            }
            else
            {
                var servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.SearchPriceForCustomer(customerName1, siteMAD, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2), "20", datasheetNameMAD);

                var serviceGeneralInformationsPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameNormalMAD);

            string serviceNormalMAD = servicePage.GetFirstServiceName();

            // -------------------------------------------- service vacuum --------------------------------------------------
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameVacuum);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceNameVacuum, null, null, serviceCategorie2, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerName2, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2), "10", datasheetName2);
                servicePage = pricePage.BackToList();
            }
            else
            {
                var servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.SearchPriceForCustomer(customerName2, siteACE, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2), "10", datasheetName2);

                var serviceGeneralInformationsPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameVacuum);

            string serviceVacuum = servicePage.GetFirstServiceName();

            // -------------------------------------------- service flight valide --------------------------------------------------

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameFlightValide);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceNameFlightValide, null, null, serviceCategorie3, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerName3, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2), "30", datasheetName3);
                servicePage = pricePage.BackToList();
            }
            else
            {
                var servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.SearchPriceForCustomer(customerName3, siteACE, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2), "30", datasheetName3);

                var serviceGeneralInformationsPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameFlightValide);

            string serviceFlightValide = servicePage.GetFirstServiceName();

            // -------------------------------------------- service sans Datasheet --------------------------------------------------
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameSansDatasheet);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceNameSansDatasheet, null, null, serviceCategorie1, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerName1, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.SearchPriceForCustomer(customerName1, siteACE, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));

                var serviceGeneralInformationsPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameSansDatasheet);

            string serviceSansDatasheet = servicePage.GetFirstServiceName();

            // -------------------------------------------------- Assert -------------------------------------------------
            Assert.IsTrue(serviceNormal.Contains(serviceNameNormal), "Le service " + serviceNameNormal + " n'existe pas.");
            Assert.IsTrue(serviceNormalMAD.Contains(serviceNameNormalMAD), "Le service " + serviceNameNormalMAD + " n'existe pas.");
            Assert.IsTrue(serviceVacuum.Contains(serviceNameVacuum), "Le service " + serviceNameVacuum + " n'existe pas.");
            Assert.IsTrue(serviceFlightValide.Contains(serviceNameFlightValide), "Le service " + serviceNameFlightValide + " n'existe pas.");
            Assert.IsTrue(serviceSansDatasheet.Contains(serviceNameSansDatasheet), "Le service " + serviceNameSansDatasheet + " n'existe pas.");
        }

        [TestMethod]
        [Priority(6)]
        [Timeout(_timeout)]

        public void PU_NEEDS_CreateFlight()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string customerName2 = TestContext.Properties["Needs_Customer2"].ToString();
            string customerName3 = TestContext.Properties["Needs_Customer3"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameNormalMAD = TestContext.Properties["Needs_ServiceNormalMAD"].ToString();
            string serviceNameVacuum = TestContext.Properties["Needs_ServiceVacuum"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();
            string serviceNameSansDatasheet = TestContext.Properties["Needs_ServiceSansDatasheet"].ToString();
            string flightNameNormal = "" + TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameNormalMAD = "" + TestContext.Properties["Needs_FlightNormalMAD"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameVacuum = "" + TestContext.Properties["Needs_FlightVacuum"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = "" + TestContext.Properties["Needs_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameSansDatasheet = "" + TestContext.Properties["Needs_FlightSansDatasheet"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string service1;
            string serviceMAD;
            string service2;
            string service3;
            string service4;

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var flightPage = homePage.GoToFlights_FlightPage();

            // -------------------------------------------- flight normal --------------------------------------------------
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameNormal);

            if (flightPage.CheckTotalNumber() == 0)
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNameNormal, customerName1, aircraft, siteACE, siteBCN);

                //Edit the first flight
                var editPage = flightPage.EditFirstFlight(flightNameNormal);
                editPage.AddGuestType(guestType1);
                editPage.AddService(serviceNameNormal);
                service1 = editPage.GetServiceName();
                editPage.CloseViewDetails();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameNormal);
            }
            else
            {
                var editPage = flightPage.EditFirstFlight(flightNameNormal);

                if (editPage.GetGuestType().Equals("") || !editPage.GetGuestType().Equals(guestType1))
                {
                    editPage.AddGuestType(guestType1);
                    editPage.AddService(serviceNameNormal);
                }

                if (editPage.GetServiceName().Equals("") || !editPage.GetServiceName().Equals(serviceNameNormal))
                {
                    editPage.AddService(serviceNameNormal);
                }

                service1 = editPage.GetServiceName();

                editPage.CloseViewDetails();
            }
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameNormal);

            string flightNormal = flightPage.GetFirstFlightNumber();

            // -------------------------------------------- flight normal MAD --------------------------------------------------
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteMAD);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameNormalMAD);

            if (flightPage.CheckTotalNumber() == 0)
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNameNormalMAD, customerName1, aircraft, siteMAD, siteBCN);

                //Edit the first flight
                var editPage = flightPage.EditFirstFlight(flightNameNormalMAD);
                editPage.AddGuestType(guestType1);
                editPage.AddService(serviceNameNormalMAD);
                serviceMAD = editPage.GetServiceName();
                editPage.CloseViewDetails();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameNormalMAD);
            }
            else
            {
                var editPage = flightPage.EditFirstFlight(flightNameNormalMAD);

                if (editPage.GetGuestType().Equals("") || !editPage.GetGuestType().Equals(guestType1))
                {
                    editPage.AddGuestType(guestType1);
                    editPage.AddService(serviceNameNormalMAD);
                }

                if (editPage.GetServiceName().Equals("") || !editPage.GetServiceName().Equals(serviceNameNormalMAD))
                {
                    editPage.AddService(serviceNameNormalMAD);
                }

                serviceMAD = editPage.GetServiceName();

                editPage.CloseViewDetails();
            }

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameNormalMAD);
            string flightNormalMAD = flightPage.GetFirstFlightNumber();

            // -------------------------------------------- flight vacuum --------------------------------------------------
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameVacuum);

            if (flightPage.CheckTotalNumber() == 0)
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNameVacuum, customerName2, aircraft, siteACE, siteBCN);

                //Edit the first flight
                var editPage = flightPage.EditFirstFlight(flightNameVacuum);
                editPage.AddGuestType(guestType2);
                editPage.AddService(serviceNameVacuum);
                service2 = editPage.GetServiceName();
                editPage.CloseViewDetails();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameVacuum);
            }
            else
            {
                var editPage = flightPage.EditFirstFlight(flightNameVacuum);

                if (editPage.GetGuestType().Equals("") || !editPage.GetGuestType().Equals(guestType2))
                {
                    editPage.AddGuestType(guestType2);
                    editPage.AddService(serviceNameVacuum);
                }

                if (editPage.GetServiceName().Equals("") || !editPage.GetServiceName().Equals(serviceNameVacuum))
                {
                    editPage.AddService(serviceNameVacuum);
                }

                service2 = editPage.GetServiceName();

                editPage.CloseViewDetails();
            }

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameVacuum);
            string flightVacuum = flightPage.GetFirstFlightNumber();

            // -------------------------------------------- flight validé --------------------------------------------------
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameValide);

            if (flightPage.CheckTotalNumber() == 0)
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNameValide, customerName3, aircraft, siteACE, siteBCN);

                //Edit the first flight
                var editPage = flightPage.EditFirstFlight(flightNameValide);
                editPage.AddGuestType(guestType2);
                editPage.AddService(serviceNameFlightValide);
                service3 = editPage.GetServiceName();
                editPage.CloseViewDetails();

                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameValide);
                flightPage.SetNewState("V");
                WebDriver.Navigate().Refresh();
            }
            else
            {
                var editPage = flightPage.EditFirstFlight(flightNameValide);

                bool isValid = editPage.GetFlightStatus("V");

                if (editPage.GetGuestType().Equals("") || !editPage.GetGuestType().Equals(guestType2))
                {
                    editPage.AddGuestType(guestType2);
                    editPage.AddService(serviceNameFlightValide);
                }

                if (editPage.GetServiceName(isValid).Equals("") || !editPage.GetServiceName(isValid).Equals(serviceNameFlightValide))
                {
                    editPage.AddService(serviceNameFlightValide);
                }

                service3 = editPage.GetServiceName(isValid);

                editPage.CloseViewDetails();

                if (!isValid)
                {
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameValide);
                    flightPage.SetNewState("V");
                    WebDriver.Navigate().Refresh();
                }
            }

            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameValide);
            string flightValide = flightPage.GetFirstFlightNumber();
            bool isStatusOK = flightPage.GetFirstFlightStatus("V");

            // -------------------------------------------- flight sans datasheet --------------------------------------------------
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameSansDatasheet);

            if (flightPage.CheckTotalNumber() == 0)
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNameSansDatasheet, customerName1, aircraft, siteACE, siteBCN);

                //Edit the first flight
                var editPage = flightPage.EditFirstFlight(flightNameSansDatasheet);
                editPage.AddGuestType(guestType2);
                editPage.AddService(serviceNameSansDatasheet);
                service4 = editPage.GetServiceName();
                editPage.CloseViewDetails();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameSansDatasheet);
            }
            else
            {
                var editPage = flightPage.EditFirstFlight(flightNameSansDatasheet);

                if (editPage.GetGuestType().Equals("") || !editPage.GetGuestType().Equals(guestType2))
                {
                    editPage.AddGuestType(guestType2);
                    editPage.AddService(serviceNameSansDatasheet);
                }

                if (editPage.GetServiceName().Equals("") || !editPage.GetServiceName().Equals(serviceNameSansDatasheet))
                {
                    editPage.AddService(serviceNameSansDatasheet);
                }

                service4 = editPage.GetServiceName();

                editPage.CloseViewDetails();
            }

            string flightSansDatasheet = flightPage.GetFirstFlightNumber();

            // -------------------------------------------------- Assert flights -------------------------------------------------
            Assert.AreEqual(flightNameNormal, flightNormal, "Le flight " + flightNameNormal + " n'existe pas.");
            Assert.AreEqual(serviceNameNormal, service1, "Le flight " + flightNameNormal + " n'est pas associé au service " + serviceNameNormal + ".");
            Assert.AreEqual(flightNameNormalMAD, flightNormalMAD, "Le flight " + flightNameNormalMAD + " n'existe pas.");
            Assert.AreEqual(serviceNameNormalMAD, serviceMAD, "Le flight " + flightNameNormalMAD + " n'est pas associé au service " + serviceNameNormalMAD + ".");
            Assert.AreEqual(flightNameVacuum, flightVacuum, "Le flight " + flightNameVacuum + " n'existe pas.");
            Assert.AreEqual(serviceNameVacuum, service2, "Le flight " + flightNameVacuum + " n'est pas associé au service " + serviceNameVacuum + ".");
            Assert.AreEqual(flightNameValide, flightValide, "Le flight " + flightNameValide + " n'existe pas.");
            Assert.IsTrue(isStatusOK, "Le flight " + flightNameValide + " n'est pas validé.");
            Assert.AreEqual(serviceNameFlightValide, service3, "Le flight " + flightNameValide + " n'est pas associé au service " + serviceNameFlightValide + ".");
            Assert.AreEqual(flightNameSansDatasheet, flightSansDatasheet, "Le flight " + flightNameSansDatasheet + " n'existe pas.");
            Assert.AreEqual(serviceNameSansDatasheet, service4, "Le flight " + flightNameSansDatasheet + " n'est pas associé au service " + serviceNameSansDatasheet + ".");

            // ------------------------------------------------- Schedule flight ------------------------------------------
            var schedulePage = homePage.GoToFlights_FlightSelectionPage();
            schedulePage.ResetFilter();
            schedulePage.PageSize("100");
            schedulePage.Filter(SchedulePage.FilterType.Site, siteACE);

            schedulePage.UnfoldAll();
            schedulePage.SetFlightProduced(flightNameNormal, true);
            bool isFlight1Produced = schedulePage.IsFlightProducedToday(flightNameNormal);

            schedulePage.SetFlightProduced(flightNameVacuum, true);
            bool isFlight2Produced = schedulePage.IsFlightProducedToday(flightNameVacuum);

            schedulePage.SetFlightProduced(flightNameValide, true);
            bool isFlight3Produced = schedulePage.IsFlightProducedToday(flightNameValide);

            schedulePage.SetFlightProduced(flightNameSansDatasheet, true);
            bool isFlight4Produced = schedulePage.IsFlightProducedToday(flightNameSansDatasheet);

            schedulePage.Filter(SchedulePage.FilterType.Site, siteMAD);
            schedulePage.UnfoldAll();
            schedulePage.SetFlightProduced(flightNameNormalMAD, true);
            bool isFlightMADProduced = schedulePage.IsFlightProducedToday(flightNameNormalMAD);

            Assert.IsTrue(isFlight1Produced, "Le flight " + flightNameNormal + " n'a pas de production pour D.");
            Assert.IsTrue(isFlight2Produced, "Le flight " + flightNameVacuum + " n'a pas de production pour D.");
            Assert.IsTrue(isFlight3Produced, "Le flight " + flightNameValide + " n'a pas de production pour D.");
            Assert.IsTrue(isFlight4Produced, "Le flight " + flightNameSansDatasheet + " n'a pas de production pour D.");
            Assert.IsTrue(isFlightMADProduced, "Le flight " + flightNameNormalMAD + " n'a pas de production pour D.");
        }

        // _____________________________________________ Filtres et page Filter And Favorite__________________________________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_Site()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameNormalMAD = TestContext.Properties["Needs_ServiceNormalMAD"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            // Vérification des services pour les différents sites
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + "n'est pas présent pour le site " + siteACE);
            Assert.IsFalse(qtyAjustementPage.IsServicePresent(serviceNameNormalMAD), "Le service " + serviceNameNormalMAD + "est présent pour le site " + siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteMAD);
            Assert.IsFalse(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + "est présent pour le site " + siteMAD);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormalMAD), "Le service " + serviceNameNormal + "est  présent pour le site " + siteMAD);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_Site_FaF()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string serviceNameNormalMAD = TestContext.Properties["Needs_ServiceNormalMAD"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();

            // On filtre les données par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, site);
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            string siteSelected = qtyAjustementPage.GetSite();

            Assert.AreEqual(site, siteSelected, String.Format(MessageErreur.FILTRE_ERRONE, "Site FaF"));
            Assert.IsFalse(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + "est présent pour le site " + site);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormalMAD), "Le service " + serviceNameNormal + "n'est pas présent pour le site " + site);
        }

        [TestMethod]

        [Timeout(_timeout)]
        public void PU_NEEDS_Filter_SortByCustomer()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameVacuum = TestContext.Properties["Needs_ServiceVacuum"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + "n'est pas présent pour le site " + siteACE);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameVacuum), "Le service " + serviceNameVacuum + "est présent pour le site " + siteACE);
            // Sort by customer
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.SortBy, "Customer");
            bool isSortedByCustomer = qtyAjustementPage.IsSortedByCustomer();

            //assert
            Assert.IsTrue(isSortedByCustomer, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by customers"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_FromTo()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameNormalAvant = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.AddDays(-2).ToShortDateString() + "testDate";

            DateTime dateFrom = DateUtils.Now.AddDays(-2);
            DateTime dateTo = DateUtils.Now;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            qtyAjustementPage.PageSize("100");
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.DateFrom, dateFrom);
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.DateTo, dateTo);
            bool isServiceNormalAvant = qtyAjustementPage.IsServicePresent(serviceNameNormal);

            if (!isServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);
            }

            if (!isServiceNormalAvant)
            {
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
                flightPage.SetDateState(dateFrom);

                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNameNormalAvant, customerName1, aircraft, siteACE, siteBCN);
                var editPage = flightPage.EditFirstFlight(flightNameNormalAvant);
                editPage.AddGuestType(guestType1);
                editPage.AddService(serviceNameNormal);
                editPage.CloseViewDetails();

                var schedulePage = homePage.GoToFlights_FlightSelectionPage();
                schedulePage.ResetFilter();
                schedulePage.Filter(SchedulePage.FilterType.Site, siteACE);
                schedulePage.Filter(SchedulePage.FilterType.DateFrom, dateFrom);
                schedulePage.Filter(SchedulePage.FilterType.DateTo, dateFrom);

                schedulePage.UnfoldAll();
                schedulePage.SetFlightProduced(flightNameNormalAvant, true);
            }

            if (!isServiceNormal || !isServiceNormalAvant)
            {
                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            }

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.DateFrom, dateFrom);
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.DateTo, dateTo);

            // Page qty adjustement
            bool isDateFromOkQtyAdj = qtyAjustementPage.VerifyDateFrom(dateFrom);
            bool isDateToOkQtyAdj = qtyAjustementPage.VerifyDateTo(dateTo);

            // Page results
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);

            bool isDateFromOkResult = resultPage.VerifyDateFrom(dateFrom.Date);
            bool isDateToOkResult = resultPage.VerifyDateTo(dateTo.Date);

            // Assert
            Assert.IsTrue(isDateFromOkQtyAdj, String.Format(MessageErreur.FILTRE_ERRONE, "From Qty adjustement"));
            Assert.IsTrue(isDateToOkQtyAdj, String.Format(MessageErreur.FILTRE_ERRONE, "To Qty adjustement"));
            Assert.IsTrue(isDateFromOkResult, String.Format(MessageErreur.FILTRE_ERRONE, "From Results"));
            Assert.IsTrue(isDateToOkResult, String.Format(MessageErreur.FILTRE_ERRONE, "To Results"));
        }

        [TestMethod]

        [Timeout(_timeout)]
        public void PU_NEEDS_Filter_FromTo_FaF()
        {
            // Prepare
            DateTime dateFrom = DateUtils.Now.AddDays(-2);
            DateTime dateTo = DateUtils.Now.AddDays(+1);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();

            // On filtre par production Date
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.DateFrom, dateFrom);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.DateTo, dateTo);
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            // On récupère les dates dans la page qty adjustement
            DateTime dateFromFilter = qtyAjustementPage.GetDateFrom();
            DateTime dateToFilter = qtyAjustementPage.GetDateTo();

            Assert.AreEqual(dateFrom.Date, dateFromFilter.Date, String.Format(MessageErreur.FILTRE_ERRONE, "From FaF"));
            Assert.AreEqual(dateTo.Date, dateToFilter.Date, String.Format(MessageErreur.FILTRE_ERRONE, "To FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_ShowWithoutDatasheet()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameSansDatasheet = TestContext.Properties["Needs_ServiceSansDatasheet"].ToString();
            string flightNameSansDatasheet = TestContext.Properties["Needs_FlightSansDatasheet"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            DateTime dateFrom = DateTime.ParseExact("24/09/2024", "dd/MM/yyyy", CultureInfo.InvariantCulture);
            DateTime dateTo = DateTime.ParseExact("24/09/2024", "dd/MM/yyyy", CultureInfo.InvariantCulture);

            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.DateFrom, dateFrom);

            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.DateTo, dateTo);

            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceSansDatasheet = qtyAjustementPage.IsServicePresent(serviceNameSansDatasheet);

            if (!isServiceSansDatasheet)
            {
                CreateFlight(homePage, siteACE, flightNameSansDatasheet, customerName1, guestType2, serviceNameSansDatasheet);

                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            }

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.ShowWithoutDatasheet, true);
            qtyAjustementPage.PageSize("100");

            // Page qty adjustement
            bool isWithoutDatasheet = qtyAjustementPage.IsWithoutDatasheet();

            // Page results
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            //assert
            Assert.IsTrue(isWithoutDatasheet, String.Format(MessageErreur.FILTRE_ERRONE, "Show services without datasheet only"));
            Assert.AreEqual(0, resultPage.CountResults(), String.Format(MessageErreur.FILTRE_ERRONE, "Show services without datasheet only"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_NEEDS_Filter_ShowValidatedFlightsOnly()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();

            //Arrange
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);


            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas présent.");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameFlightValide), "Le service " + serviceNameFlightValide + " n'est pas présent.");

            ResultPageNeeds resultPage = qtyAjustementPage.GoToResultPage();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.ShowValidateFlight, true);

            resultPage.GoToQtyAdjustementPage();
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameFlightValide), "Le service " + serviceNameFlightValide + " n'est pas présent.");
            Assert.IsFalse(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas présent.");

            // Page results
            resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            HashSet<string> flights = resultPage.GetFlights();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.SetDateState(DateUtils.Now.AddDays(1));
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);


            foreach (string flightName in flights)
            {
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightName);
                Assert.IsTrue(flightPage.GetFirstFlightStatus("V"), String.Format(MessageErreur.FILTRE_ERRONE, "Show validated flight only"));
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_ShowHiddenArticle()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();

            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string customerName3 = TestContext.Properties["Needs_Customer3"].ToString();

            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();

            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = TestContext.Properties["Needs_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            if (!isServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);
            }

            if (!isServiceValide)
            {
                CreateFlight(homePage, siteACE, flightNameValide, customerName3, guestType2, serviceNameFlightValide, true);
            }

            if (!isServiceNormal || !isServiceValide)
            {
                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            }

            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas présent.");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameFlightValide), "Le service " + serviceNameFlightValide + " n'est pas présent.");

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            resultPage.UnfoldAll();
            string firstItemName = resultPage.GetFirstItemName();

            var items = resultPage.GetItemNames();
            int nbItemInit = resultPage.CountItems();

            Assert.IsTrue(items.Contains(firstItemName), "La liste initiale ne contient pas l'item " + firstItemName);

            // On cache le premier article (icone d'action, on hide l'item et non icone d'état)
            try
            {


                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

                resultPage = qtyAjustementPage.GoToResultPage();
                resultPage.FilterRawBy(RAW_MAT_GROUP);
                resultPage.UnfoldAll();
                resultPage.HideFirstArticle();
                items = resultPage.GetItemNames();
                int nbItemFinal = resultPage.CountItems();

                //assert
                Assert.IsFalse(items.Contains(firstItemName), "L'item " + firstItemName + " n'a pas été caché.");
                Assert.AreNotEqual(nbItemInit, nbItemFinal, "le nombre d'items visible n'a pas été modifié.");


                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.ShowHiddenArticles, true);
                resultPage.UnfoldAll();
                resultPage.UnfoldAll();

                items = resultPage.GetItemNames();
                Assert.IsTrue(items.Contains(firstItemName), String.Format(MessageErreur.FILTRE_ERRONE, "Show hidden article"));
                Assert.AreEqual(nbItemInit, resultPage.CountItems(), "Le nombre d'items visibles n'a pas été retabli malgré l'utilisation du filtre.");
            }
            // On fait réapparaitre le premier item
            finally
            {
                resultPage.HideFirstArticle();

                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

                resultPage = qtyAjustementPage.GoToResultPage();
                resultPage.FilterRawBy(RAW_MAT_GROUP);
                resultPage.UnfoldAll();

                items = resultPage.GetItemNames();

                //assert
                Assert.IsTrue(items.Contains(firstItemName), "L'item " + firstItemName + " n'a pas été retabli.");
                Assert.AreEqual(nbItemInit, resultPage.CountItems(), "le nombre d'items visible n'a pas été rétabli.");
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_ShowNormalAndVacuumProduction()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string recipeCookingMode = TestContext.Properties["Prodman_Needs_CookingMode1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Prodman_Needs_CookingMode2"].ToString();


            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameVacuum = TestContext.Properties["Needs_ServiceVacuum"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceVacuum = qtyAjustementPage.IsServicePresent(serviceNameVacuum);
            Assert.IsTrue(isServiceNormal, "Le service " + serviceNameNormal + " n'est pas visible.");
            Assert.IsTrue(isServiceVacuum, "Le service " + serviceNameVacuum + " n'est pas visible.");

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.ShowNormalAndVacuumProd, true);

            var resultPage = qtyAjustementPage.GoToResultPage();

            // ___________________ RAW MAT BY GROUP ______________________________
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            var cookingModesGroup = resultPage.GetCookingModes();
            bool isProdOKGroup = (cookingModesGroup.Count >= 1 && cookingModesGroup.Contains(recipeCookingMode1) && cookingModesGroup.Contains(recipeCookingMode1));

            // ___________________ RAW MAT WORKSHOP ______________________________
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            var cookingModesWorkshop = resultPage.GetCookingModesForWorkShop();
            bool isProdOKWorkshop = (cookingModesWorkshop.Count >= 1 && cookingModesWorkshop.Contains(recipeCookingMode) && cookingModesWorkshop.Contains(recipeCookingMode1));

            // ___________________ RAW MAT SUPPLIER ______________________________
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            var cookingModesSupplier = resultPage.GetCookingModes();
            bool isProdOKSupplier = (cookingModesSupplier.Count >= 1 && cookingModesSupplier.Contains(recipeCookingMode));

            // ___________________ RAW MAT RECIPE ______________________________
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            var cookingModesRecipe = resultPage.GetCookingModesForWorkShop();
            bool isProdOKRecipe = (cookingModesRecipe.Count >= 1 && cookingModesRecipe.Contains(recipeCookingMode) && cookingModesRecipe.Contains(recipeCookingMode1));

            // Assert
            Assert.IsTrue(isProdOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal and vacuum production Raw mat by group"));
            Assert.IsTrue(isProdOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal and vacuum production Raw mat by Workshop"));
            Assert.IsTrue(isProdOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal and vacuum production Raw mat by Supplier"));
            Assert.IsTrue(isProdOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal and vacuum production Raw mat by Recipe"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_NEEDS_Filter_ShowNormalAndVacuumProduction_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            // Filtre par "show normal and vacuum" et par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.ShowNormalAndVacuumProd, true);
            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            bool isNormalAndVacuumProduction = qtyAjustementPage.IsNormalAndVacuumProduction();
            //assert
            Assert.IsTrue(isNormalAndVacuumProduction, string.Format(MessageErreur.FILTRE_ERRONE, "Show normal and vacuum production FaF"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_ShowVacuumProductionOnly()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string recipeCookingMode = TestContext.Properties["Prodman_Needs_CookingMode2"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameVacuum = TestContext.Properties["Needs_ServiceVacuum"].ToString();


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameVacuum), "Le service " + serviceNameVacuum + " n'est pas visible.");

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.ShowVacuumProd, true);

            var resultPage = qtyAjustementPage.GoToResultPage();

            // ___________________ RAW MAT BY GROUP ______________________________
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            var cookingModesGroup = resultPage.GetCookingModes();
            bool isVacuumOKGroup = (cookingModesGroup.Count == 1 && cookingModesGroup.Contains(recipeCookingMode));

            // ___________________ RAW MAT WORKSHOP ______________________________
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            var cookingModesWorkshop = resultPage.GetCookingModesForWorkShop();
            bool isVacuumOKWorkshop = (cookingModesWorkshop.Count == 1 && cookingModesWorkshop.Contains(recipeCookingMode));

            // ___________________ RAW MAT SUPPLIER ______________________________
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            var cookingModesSupplier = resultPage.GetCookingModes();
            bool isVacuumOKSupplier = (cookingModesSupplier.Count == 1 && cookingModesSupplier.Contains(recipeCookingMode));

            // ___________________ RAW MAT RECIPE ______________________________
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            var cookingModesRecipe = resultPage.GetCookingModesForWorkShop();
            bool isVacuumOKRecipe = (cookingModesRecipe.Count == 1 && cookingModesRecipe.Contains(recipeCookingMode));

            // Assert
            Assert.IsTrue(isVacuumOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Show vacuum production only Raw mat by group"));
            Assert.IsTrue(isVacuumOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Show vacuum production only Raw mat by Workshop"));
            Assert.IsTrue(isVacuumOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Show vacuum production only Raw mat by Supplier"));
            Assert.IsTrue(isVacuumOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Show vacuum production only Raw mat by Recipe"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_ShowVacuumProductionOnly_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "show vacuum only" et par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.ShowVacuumProd, true);

            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            bool isVacuumProduction = qtyAjustementPage.IsVacuumProduction();

            //assert
            Assert.IsTrue(isVacuumProduction, String.Format(MessageErreur.FILTRE_ERRONE, "Show vacuum production only FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_ShowNormalProductionOnly()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string recipeCookingMode = TestContext.Properties["Prodman_Needs_CookingMode2"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameVacuum = TestContext.Properties["Needs_ServiceVacuum"].ToString();


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameVacuum), "Le service " + serviceNameVacuum + " n'est pas visible.");

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.ShowNormalProd, true);

            var resultPage = qtyAjustementPage.GoToResultPage();

            // ___________________ RAW MAT BY GROUP ______________________________
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            var cookingModesGroup = resultPage.GetCookingModes();
            bool isNormalOKGroup = (cookingModesGroup.Count != 0 && !cookingModesGroup.Contains(recipeCookingMode));

            // ___________________ RAW MAT WORKSHOP ______________________________
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            var cookingModesWorkshop = resultPage.GetCookingModesForWorkShop();
            bool isNormalOKWorkshop = (cookingModesWorkshop.Count != 0 && !cookingModesWorkshop.Contains(recipeCookingMode));

            // ___________________ RAW MAT SUPPLIER ______________________________
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            var cookingModesSupplier = resultPage.GetCookingModes();
            bool isNormalOKSupplier = (cookingModesSupplier.Count != 0 && !cookingModesSupplier.Contains(recipeCookingMode));

            // ___________________ RAW MAT RECIPE ______________________________
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            var cookingModesRecipe = resultPage.GetCookingModesForWorkShop();
            bool isNormalOKRecipe = (cookingModesRecipe.Count != 0 && !cookingModesRecipe.Contains(recipeCookingMode));

            // Assert
            Assert.IsTrue(isNormalOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal production only Raw mat by group"));
            Assert.IsTrue(isNormalOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal production only Raw mat by Workshop"));
            Assert.IsTrue(isNormalOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal production only Raw mat by Supplier"));
            Assert.IsTrue(isNormalOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal production only Raw mat by Recipe"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_ShowNormalProductionOnly_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "show normal only" et par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.ShowNormalProd, true);

            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            bool isNormalProduction = qtyAjustementPage.IsNormalProduction();
            //assert
            Assert.IsTrue(isNormalProduction, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal production only FaF"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_NEEDS_Filter_Workshops()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameFlightValide), "Le service " + serviceNameFlightValide + " n'est pas visible.");


            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.GroupBy(GROUP_BY_WORKSHOP);
            string workshop = resultPage.GetFirstItemGroup();

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Workshops, workshop);

            // ___________________ RAW MAT BY GROUP ______________________________
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            string workshopReturned = resultPage.GetFirstItemGroup();
            resultPage.UnfoldAll();
            Assert.AreEqual(workshop, workshopReturned, "Le nom du workshop ne correspond pas à celui filtré dans 'Raw mat by group'.");
            bool isWorkshopOKGroup = resultPage.VerifyWorkshop(workshop);
            Assert.IsTrue(isWorkshopOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Workshop Raw mat by group"));
            // ___________________ RAW MAT WORKSHOP ______________________________
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isWorkshopOKWorkshop = resultPage.VerifyWorkshopForRawMatWorkShop(workshop);
            Assert.IsTrue(isWorkshopOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Workshop Raw mat by Workshop"));
            // ___________________ RAW MAT SUPPLIER ______________________________
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isWorkshopOKSupplier = resultPage.VerifyWorkshop(workshop);
            Assert.IsTrue(isWorkshopOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Workshop Raw mat by Supplier"));
            // ___________________ RAW MAT RECIPE ______________________________
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isWorkshopOKRecipe = resultPage.VerifyWorkshopForRecipe(workshop);
            Assert.IsTrue(isWorkshopOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Workshop Raw mat by Recipe"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_NEEDS_Filter_Workshops_FaF()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Prodman_Needs_Workshop1"].ToString();
            var homePage = LogInAsAdmin();
            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Workshops, recipeWorkshop1);
            string selectedRcipeTypeNbr = filterAndFavorite.CheckNbrWorkshopsSelected();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            string selectedRcipeTypeNbrFromResult = filterAndFavorite.CheckNbrWorkshopsSelected();
            string selectedWorkshop = qtyAjustementPage.GetWorkshop();
            Assert.AreEqual(recipeWorkshop1, selectedWorkshop, string.Format(MessageErreur.FILTRE_ERRONE, "Workshop FaF"));
            Assert.AreEqual(selectedRcipeTypeNbr, selectedRcipeTypeNbrFromResult, string.Format(MessageErreur.FILTRE_ERRONE, "Workshop FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_Customers()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string customerName3 = TestContext.Properties["Needs_Customer3"].ToString();

            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNameFlightValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            Assert.IsTrue(isServiceNameFlightValide, "Le service " + serviceNameFlightValide + "n'est pas présent pour le site " + siteACE);

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Customers, customerName3);
            qtyAjustementPage.PageSize("100");

            // Page Qty ajustement
            bool isCustomerOKQtyAjus = qtyAjustementPage.VerifyCustomer(customerName3);

            // Page results
            // ___________________ RAW MAT BY GROUP ______________________________
            ResultPageNeeds resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.ShowHidden();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.GroupBy(GROUP_BY_CUSTOMER);
            resultPage.WaitPageLoading();
            resultPage.UnfoldAll();

            bool customerNameGroup = resultPage.GetFirstItemGroup().Equals(customerName3) && (resultPage.CountResults() == 1);
            bool isCustomerOKGroup = resultPage.VerifyCustomerRawMat(customerName3);

            // ___________________ RAW MAT WORKSHOP ______________________________
         resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();
            bool isCustomerOKWorkshop = resultPage.VerifyCustomerForRawMatWorkShop(customerName3);

            // ___________________ RAW MAT SUPPLIER ______________________________
           resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isCustomerOKSupplier = resultPage.VerifyCustomerRawMat(customerName3);

            // ___________________ RAW MAT RECIPE ______________________________
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool customerNameRecipe = resultPage.GetSubGroupName().Equals(customerName3) && (resultPage.CountSubResults() == 1);
            bool isCustomerOKRecipe = resultPage.VerifyCustomerForRawMatWorkShop(customerName3);

            // ___________________ RAW MAT CUSTOMER ______________________________
            resultPage.FilterRawBy(RAW_MAT_CUSTOMER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool customerName = resultPage.GetSubGroupName().Equals(customerName3) && (resultPage.CountSubResults() == 1);

            // Assert
            Assert.IsTrue(isCustomerOKQtyAjus, String.Format(MessageErreur.FILTRE_ERRONE, "Customers qty ajustement"));
            Assert.IsTrue(customerNameGroup, "Le nom du customer filtré n'est pas correct dans le Raw mat by group.");
            Assert.IsTrue(isCustomerOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Customers raw mat by group"));
            Assert.IsTrue(isCustomerOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Customers raw mat by workshop"));
            Assert.IsTrue(isCustomerOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Customers raw mat by supplier"));
            Assert.IsTrue(customerNameRecipe, "Le nom du customer filtré n'est pas correct dans le Raw mat by recipe.");
            Assert.IsTrue(isCustomerOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Customers raw mat by recipe"));
            Assert.IsTrue(customerName, String.Format(MessageErreur.FILTRE_ERRONE, "Customers raw mat by customer"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_Customers_FaF()
        {
            // Préparation des données de test
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            // Arrange
            var homePage = LogInAsAdmin();
            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            // Application du filtre par site et client
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Customers, customerName1);
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            // Récupération 
            List<string> filteredResults = qtyAjustementPage.GetCustomers();
            //Il existe au moins un résultat correspondant aux filtres
            bool isSelected = filteredResults.Count > 0;
            Assert.IsTrue(isSelected, "Aucun résultat trouvé après application des filtres.");
            //Tous les résultats contiennent le nom du client attendu
            bool isFiltred = filteredResults.All(result => result.Contains(customerName1));
            Assert.IsTrue(isFiltred, String.Format(MessageErreur.FILTRE_ERRONE, "Customer FaF"));
            bool isServicePresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(isServicePresent, "le service relier au customer normal n est pas afficher ");
        }
      
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_GuestType()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            bool isServiceNormalPresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValidePresent = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            Assert.IsTrue(isServiceNormalPresent == true && isServiceValidePresent == true, "un de service n est pas disponible ");
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.GuestType, guestType1);
            qtyAjustementPage.PageSize("100");

            // Page qty ajustements
            isServiceNormalPresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            isServiceValidePresent = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            Assert.IsTrue(isServiceNormalPresent == true && isServiceValidePresent == false, String.Format(MessageErreur.FILTRE_ERRONE, "Guest type qty adjustements"));
            // Page results
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_CUSTOMER);
            resultPage.UnfoldAll();
            bool returnGuestType = resultPage.VerifyGuestType(guestType1);


            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            var flights = resultPage.GetFlights();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.SetDateState(DateUtils.Now.AddDays(1));

            foreach (string flightName in flights)
            {
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightName);
                FlightDetailsPage editPage = flightPage.EditFirstFlight(flightName);
                string guestRetourne = editPage.GetGuestType();
                flightPage = editPage.CloseViewDetails();
                Assert.AreEqual(guestType1, guestRetourne, String.Format(MessageErreur.FILTRE_ERRONE, "Guest type"));
            }

            Assert.IsTrue(returnGuestType, String.Format(MessageErreur.FILTRE_ERRONE, "Guest type results"));
        }
       
        [Timeout(_timeout)]
        [TestMethod]

        public void PU_NEEDS_Filter_GuestType_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();

            // Filtre par guestType et par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.GuestType, guestType1);

            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            string selectedGuestType = qtyAjustementPage.GetGuestType();

            //assert
            Assert.AreEqual(guestType1, selectedGuestType, String.Format(MessageErreur.FILTRE_ERRONE, "Guest type FaF"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_ServiceCategorie()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceCategorie1 = TestContext.Properties["Prodman_Needs_ServiceCategory1"].ToString();    // BEBIDAS 

            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);

            Assert.IsTrue(isServiceNormal, "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.ServicesCategorie, serviceCategorie1);

            var resultPage = qtyAjustementPage.GoToResultPage();

            // ___________________ RAW MAT BY GROUP ______________________________
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            bool isCategorieOKGroup = resultPage.VerifyCategorie(serviceCategorie1);

            // ___________________ RAW MAT WORKSHOP ______________________________
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isCategorieOKWorkshop = resultPage.VerifyCategorieForRawMatWorkShop(serviceCategorie1);

            // ___________________ RAW MAT SUPPLIER ______________________________
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isCategorieOKSupplier = resultPage.VerifyCategorie(serviceCategorie1);

            // ___________________ RAW MAT RECIPE ______________________________
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

          bool isCategorieOKRecipe = resultPage.VerifyCategorieForRawMatWorkShop(serviceCategorie1);

            //assert
            Assert.IsTrue(isCategorieOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Service categorie raw mat by group"));
            Assert.IsTrue(isCategorieOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Service categorie raw mat by workshop"));
            Assert.IsTrue(isCategorieOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Service categorie raw mat by supplier"));
            Assert.IsTrue(isCategorieOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Service categorie raw mat by recipe"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_NEEDS_Filter_ServiceCategorie_FaF()
        {
            // Prepare
            string siteMAD = TestContext.Properties["Site"].ToString();
            string serviceCategorie1 = TestContext.Properties["Prodman_Needs_ServiceCategory1"].ToString(); // BEBIDAS
            //Arrange
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            // Filtre par guestType et par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteMAD);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.ServicesCategorie, serviceCategorie1);
            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            string selectedServiceCategorie = qtyAjustementPage.GetServiceCategorie();
            //assert
            Assert.AreEqual(serviceCategorie1, selectedServiceCategorie, string.Format(MessageErreur.FILTRE_ERRONE, "Service categorie FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_Service()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(isServiceNormal, "Le service " + isServiceNormal + "n'est pas présent pour le site " + siteACE);

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Services, serviceNameNormal);
            qtyAjustementPage.PageSize("100");

            // page qty adjustement
            Assert.AreEqual(1, qtyAjustementPage.CountServices(), "Il y a plusieurs services visibles malgré le filtre.");
            var returnService = qtyAjustementPage.GetServiceName();

            // Page results
            var resultPage = qtyAjustementPage.GoToResultPage();
            // ___________________ RAW MAT BY GROUP ______________________________
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            bool isServiceOKGroup = resultPage.VerifyService(serviceNameNormal);

            // ___________________ RAW MAT WORKSHOP ______________________________
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isServiceOKWorkshop = resultPage.VerifyServiceForRawMatWorkShop(serviceNameNormal);

            // ___________________ RAW MAT SUPPLIER ______________________________
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isServiceOKSupplier = resultPage.VerifyService(serviceNameNormal);

            // ___________________ RAW MAT RECIPE ______________________________
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isServiceOKRecipe = resultPage.VerifyServiceRawMat(serviceNameNormal);

            //assert
            Assert.AreEqual(returnService, serviceNameNormal, string.Format(MessageErreur.FILTRE_ERRONE, "Filter Service qty adjustement"));
            Assert.IsTrue(isServiceOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Service raw mat by group"));
            Assert.IsTrue(isServiceOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Service raw mat by workshop"));
            Assert.IsTrue(isServiceOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Service raw mat by supplier"));
            Assert.IsTrue(isServiceOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Service raw mat by recipe"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_Service_FaF()
        {
            // Prepare
            string siteMAD = TestContext.Properties["Site"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormalMAD"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();

            // Filtre par guestType et par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteMAD);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Service, serviceNameNormal);

            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            string selectedService = qtyAjustementPage.GetService();

            //assert
            Assert.AreEqual(serviceNameNormal, selectedService, String.Format(MessageErreur.FILTRE_ERRONE, "Service FaF"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_RecipeType()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            Assert.IsTrue(isServiceValide, "Le service " + serviceNameFlightValide + "n'est pas présent pour le site " + siteACE);

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.GroupbyRawMatGroup(GROUP_BY_RECIPE_TYPE);
            string recipeType = resultPage.GetFirstItemGroup();

            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.RecipeType, recipeType);

            // ___________________ RAW MAT BY GROUP ______________________________
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            string recipeTypeReturned = resultPage.GetFirstItemGroup();

            Assert.AreEqual(recipeType, recipeTypeReturned, "Le nom du recipe type ne correspond pas à celui filtré dans 'Raw mat by group'.");

            resultPage.UnfoldAll();

            bool isRecipeTypeOKGroup = resultPage.VerifyRecipeType(recipeType);
            Assert.IsTrue(isRecipeTypeOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Recipe type Raw mat by group"));

            // ___________________ RAW MAT WORKSHOP ______________________________
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isRecipeTypeOKWorkshop = resultPage.VerifyRecipeForRawMatWorkShop(recipeType);

            Assert.IsTrue(isRecipeTypeOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Recipe type Raw mat by Workshop"));

            // ___________________ RAW MAT SUPPLIER ______________________________
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isRecipeTypeOKSupplier = resultPage.VerifyRecipeType(recipeType);

            Assert.IsTrue(isRecipeTypeOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Recipe type Raw mat by Supplier"));

            // ___________________ RAW MAT RECIPE ______________________________
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isRecipeTypeOKRecipe = resultPage.VerifyRecipeForRawMatWorkShop(recipeType);


            Assert.IsTrue(isRecipeTypeOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Recipe type Raw mat by Recipe"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_RecipeType_FaF()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string recipeType1 = TestContext.Properties["Prodman_Needs_RecipeType1"].ToString();
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.RecipeType, recipeType1);
            string selectedRcipeTypeNbr = filterAndFavorite.CheckNbrRecipeTypeSelected();
            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            string selectedRcipeTypeNbrFromResult = filterAndFavorite.CheckNbrRecipeTypeSelected();
            string selectedRecipeType = qtyAjustementPage.GetRecipeType();
            Assert.AreEqual(recipeType1, selectedRecipeType, string.Format(MessageErreur.FILTRE_ERRONE, "RecipeType FaF"));
            Assert.AreEqual(selectedRcipeTypeNbr, selectedRcipeTypeNbrFromResult, string.Format(MessageErreur.FILTRE_ERRONE, "RecipeType FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_ItemGroup()
        {
            // Prepare
            string site = TestContext.Properties["SiteMAD"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormalMAD"].ToString();
            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            Assert.AreEqual("All", filterAndFavorite.CheckTAllItemGroupsSelected(), "Tous les item Groups ne sont pas sélectionnés en arrivant sur la page : " + filterAndFavorite.CheckTAllItemGroupsSelected());

            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, site);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);

            if (!isServiceNormal)
            {
                CreateFlight(homePage, site, flightNameNormal, customerName1, guestType1, serviceNameNormal);

                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, site);
            }

            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.GroupbyRawMatGroup(GROUP_BY_ITEMGROUP);

            Assert.AreNotEqual(0, resultPage.CountResults(), "Il n'y a pas de groupes d'items visibles.");

            var firstItemGroup = resultPage.GetFirstItemGroup();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.ItemGroups, firstItemGroup);

            Assert.AreEqual(1, resultPage.CountResults(), "Il plusieurs groupes visibles après application du filtre ItemGroups.");
            Assert.AreEqual(firstItemGroup, resultPage.GetFirstItemGroup(), string.Format(MessageErreur.FILTRE_ERRONE, "ItemGroups"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Filter_ItemGroup_FaF()
        {
            //Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string itemGroup1 = TestContext.Properties["Prodman_Needs_Item_Group1"].ToString();
            //Arraange
            HomePage homePage = LogInAsAdmin();
            //Act
            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.ItemGroups, itemGroup1);
            string selectedItemGroupsNbr = filterAndFavorite.CheckNbrItemGroupsSelected();
            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            string selectedItemGroupsResultsNbr = filterAndFavorite.CheckNbrItemGroupsSelectedFromResults();
            string selectedItemGroup = qtyAjustementPage.GetItemGroup();
            //Assert
            Assert.AreEqual(itemGroup1, selectedItemGroup, String.Format(MessageErreur.FILTRE_ERRONE, "ItemGroup FaF"));
            Assert.AreEqual(selectedItemGroupsResultsNbr, selectedItemGroupsNbr, String.Format(MessageErreur.FILTRE_ERRONE, "ItemGroup FaF"));
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");
        }

        // __________________________________________ Page Filter and favorite ___________________________________________________
        [Timeout(_timeout)]
        [TestMethod]


        public void PU_NEEDS_MakeFavorite()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string group1 = TestContext.Properties["Prodman_Needs_Item_Group1"].ToString();

            string favoriteName = "testNeeds";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);

            if (!isServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);

                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();

            }
            else
            {
                filterAndFavorite = qtyAjustementPage.Back();
            }

            filterAndFavorite.DeleteFavorite(favoriteName);

            // Paramétrage et création d'un favori
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Service, serviceNameNormal);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.ItemGroups, group1);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.ShowNormalProd, true);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.RawMaterialByWorkshop, true);

            filterAndFavorite.MakeFavorite(favoriteName);
            Assert.IsTrue(filterAndFavorite.IsFavoritePresent(favoriteName), "Le favori " + favoriteName + " n'a pas été créé.");

            try
            {
                filterAndFavorite.ResetFilter();

                // On active le favorite
                filterAndFavorite.SelectFavorite(favoriteName);

                Assert.IsTrue(filterAndFavorite.GetNbServiceSelected().Contains("1 of "), "Le filtre Service du favori n'a pas été conservé.");
                Assert.IsTrue(filterAndFavorite.GetNbItemGroupSelected().Contains("1 of "), "Le filtre ItemGroup du favori n'a pas été conservé.");
                Assert.IsTrue(filterAndFavorite.IsNormalProduction(), "Le filtre 'Show normal production only' du favori n'a pas été conservé.");
                Assert.IsTrue(filterAndFavorite.IsRawMatByWorkshop(), "Le filtre 'Raw mat by workshop' du favori n'a pas été conservé.");

            }
            finally
            {
                filterAndFavorite.DeleteFavorite(favoriteName);
            }
        }

        // __________________________________________ Page Quantity adjustement __________________________________________________

        [Timeout(_timeout)]
        [TestMethod]

        public void PU_NEEDS_Edit_Qty_Adjustement()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            // Si pas de services disponibles, on recréé un vol
            if (!qtyAjustementPage.IsServicePresent(serviceNameNormal))
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);

                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            }

            // On modifie la quantité du premier service visible
            var initQty = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal);
            qtyAjustementPage.SelectService(serviceNameNormal);
            qtyAjustementPage.SetQuantity((initQty + 1).ToString(), serviceNameNormal);
            WebDriver.Navigate().Refresh();

            //Assert
            qtyAjustementPage.SelectService(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsQuantityUpdated((initQty + 1).ToString(), serviceNameNormal), "La modification n'est pas affichée.");

            qtyAjustementPage.ResetAllQuantities();
        }
        [Timeout(_timeout)]
        [TestMethod]

        public void PU_NEEDS_ResetAllQuantities()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();

            //Arrange

            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);


            if (!isServiceNormal || !isServiceValide)
            {
                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            }

            // On modifie les quantités de 2 services visibles
            var initQtyNormal = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal);
            qtyAjustementPage.SelectService(serviceNameNormal);
            qtyAjustementPage.SetQuantity((initQtyNormal + 1).ToString(), serviceNameNormal);

            var initQtyValide = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameFlightValide);
            qtyAjustementPage.SelectService(serviceNameFlightValide);
            qtyAjustementPage.SetQuantity((initQtyValide + 1).ToString(), serviceNameFlightValide);

            WebDriver.Navigate().Refresh();

            // On vérifie que les quantités ont été modifiées
            var qtyServiceNameNormal = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal);
            var qtyServiceNameFlightValide = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameFlightValide);
            Assert.AreEqual((initQtyNormal + 1), qtyServiceNameNormal, "La modification n'est pas effectuée sur le service " + serviceNameNormal);
            Assert.AreEqual((initQtyValide + 1), qtyServiceNameFlightValide, "La modification n'est pas effectuée sur le service " + serviceNameFlightValide);

            // On reset les quantités
            qtyAjustementPage.ResetAllQuantities();

            //Assert
            qtyServiceNameNormal = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal);
            qtyServiceNameFlightValide = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameFlightValide);
            Assert.AreEqual(initQtyNormal, qtyServiceNameNormal, "Le reset n'a pas fonctionné sur la quantité du service " + serviceNameNormal);
            Assert.AreEqual(initQtyValide, qtyServiceNameFlightValide, "Le reset n'a pas fonctionné sur la quantité du service " + serviceNameFlightValide);
        }

        [TestMethod]

        [Timeout(_timeout)]
        public void PU_NEEDS_RazAllQuantities()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();

            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string customerName3 = TestContext.Properties["Needs_Customer3"].ToString();

            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();

            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = TestContext.Properties["Needs_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            if (!isServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);
            }

            if (!isServiceValide)
            {
                CreateFlight(homePage, siteACE, flightNameValide, customerName3, guestType2, serviceNameFlightValide, true);
            }

            if (!isServiceNormal || !isServiceValide)
            {
                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            }

            // On modifie les quantités de 2 services visibles
            var initQtyNormal = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal);

            if (initQtyNormal == 0)
            {
                qtyAjustementPage.SelectService(serviceNameNormal);
                qtyAjustementPage.SetQuantity("1", serviceNameNormal);
            }

            var initQtyValide = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameFlightValide);

            if (initQtyValide == 0)
            {
                qtyAjustementPage.SelectService(serviceNameFlightValide);
                qtyAjustementPage.SetQuantity("1", serviceNameFlightValide);
            }

            // On raz les quantités
            qtyAjustementPage.RazAllQuantities();

            //Assert
            Assert.AreEqual(0, qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal), "Le reset n'a pas fonctionné sur la quantité du service " + serviceNameNormal);
            Assert.AreEqual(0, qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameFlightValide), "Le reset n'a pas fonctionné sur la quantité du service " + serviceNameFlightValide);

            qtyAjustementPage.ResetAllQuantities();
        }
        [Timeout(_timeout)]
        [TestMethod]

        public void PU_NEEDS_NextOrPreviousDate()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            // On récupère les dates from et to actuelles
            var initFromDate = qtyAjustementPage.GetDateFrom();
            var initToDate = qtyAjustementPage.GetDateTo();

            qtyAjustementPage.ClickOnPreviousDate();
            var previousFromDate = qtyAjustementPage.GetDateFrom();
            var previousToDate = qtyAjustementPage.GetDateTo();

            Assert.AreEqual(initFromDate.AddDays(-1), previousFromDate, "Le btn 'Prev' ne revoie pas au jour précédent pour la date From.");
            Assert.AreEqual(initToDate.AddDays(-1), previousToDate, "Le btn 'Prev' ne revoie pas au jour précédent pour la date To.");

            qtyAjustementPage.ClickOnNextDate();
            var nextFromDate = qtyAjustementPage.GetDateFrom();
            var nextToDate = qtyAjustementPage.GetDateTo();

            Assert.AreEqual(initFromDate, nextFromDate, "Le btn 'Next' ne revoie pas au jour suivant pour la date From.");
            Assert.AreEqual(initToDate, nextToDate, "Le btn 'Next' ne revoie pas au jour suivant pour la date To.");
        }

        // ________________________________________________ Page Results ____________________________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_RawMatByGroup()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);

            if (!isServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);

                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            }

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            // On récupère les couples groupe/items de la page
            Dictionary<string, List<string>> mapGroupeItems = resultPage.GetGroupAndItemsAssociated();

            var itemPage = homePage.GoToPurchasing_ItemPage();

            foreach (var entry in mapGroupeItems)
            {
                foreach (string itemName in entry.Value)
                {
                    if (!String.IsNullOrEmpty(itemName))
                    {
                        itemPage.ResetFilter();
                        itemPage.Filter(ItemPage.FilterType.Search, itemName);
                        var itemGeneralInformationPage = itemPage.ClickOnFirstItem();

                        Assert.AreEqual(entry.Key, itemGeneralInformationPage.GetGroupName(), "Le groupe " + entry.Key + " associé à l'item " + itemName + " dans la page Raw mat by group n'est pas correct.");
                        itemPage = itemGeneralInformationPage.BackToList();
                    }
                }
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_GroupBy_RawMatByGroup()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange

            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Services, serviceNameNormal);

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Group by ItemGroup
            resultPage.GroupbyRawMatGroup(GROUP_BY_ITEMGROUP);
            resultPage.UnfoldAll();

            Dictionary<string, List<string>> mapGroupeItems = resultPage.GetGroupAndItemsAssociated();

            // Group by workshop
            resultPage.GroupBy(GROUP_BY_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool resultWorkshop = resultPage.GetGroupAndRecipeAssociated("Workshop");

            // Group by Customer
            resultPage.GroupBy(GROUP_BY_CUSTOMER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            var customerName = resultPage.GetFirstItemGroup();
            bool isCustomerOKGroup = resultPage.VerifyServiceAndCustomer(serviceNameNormal, customerName);

            // Group by Recipe type
            resultPage.GroupbyRawMatGroup(GROUP_BY_RECIPE_TYPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool resultRecipeType = resultPage.GetGroupAndRecipeAssociated("RecipeType");

            // Assert des résultats
            // Item Group
            var itemPage = homePage.GoToPurchasing_ItemPage();

            foreach (var entry in mapGroupeItems)
            {
                foreach (string itemName in entry.Value)
                {
                    if (!String.IsNullOrEmpty(itemName))
                    {
                        itemPage.ResetFilter();
                        itemPage.Filter(ItemPage.FilterType.Search, itemName);
                        var itemGeneralInformationPage = itemPage.ClickOnFirstItem();

                        Assert.AreEqual(entry.Key, itemGeneralInformationPage.GetGroupName(), "Le groupe " + entry.Key + " associé à l'item " + itemName + " dans la page Raw mat by group n'est pas correct.");
                        itemPage = itemGeneralInformationPage.BackToList();
                    }
                }
            }

            // Workshop
            Assert.IsTrue(resultWorkshop, "Le workshop affiché dans l'onglet Raw mat by group -- Workshop ne correspond pas à celui de la recette associée au service " + serviceNameNormal);

            // Customer
            Assert.IsTrue(isCustomerOKGroup, "Le customer affiché dans l'onglet Raw mat by group -- Customer ne correspond pas à celui du service " + serviceNameNormal);

            // Recipe type
            Assert.IsTrue(resultRecipeType, "Le recipeType affiché dans l'onglet Raw mat by group -- Recipe type ne correspond pas à celui de la recette associée au service " + serviceNameNormal);
        }
        [Timeout(_timeout)]
        [TestMethod]

        public void PU_NEEDS_RawMatByGroup_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "Raw mat by group" et par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);

            // Group by ItemGroup
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.RawMaterialByGroup, true);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.GroupBy, GROUP_BY_ITEMGROUP);

            var resultsPage = filterAndFavorite.DoneToResults();
            bool rawMatByGroupSelected = resultsPage.GetRawMatByPage(RAW_MAT_GROUP);
            bool groupByItemGroupSelected = resultsPage.GetGroupByPage(GROUP_BY_ITEMGROUP);
            Assert.IsTrue(rawMatByGroupSelected && groupByItemGroupSelected, String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by group-- ItemGroup FaF"));

            // Group by Workshop
            filterAndFavorite = resultsPage.Back();
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.RawMaterialByGroup, true);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.GroupBy, GROUP_BY_WORKSHOP);

            resultsPage = filterAndFavorite.DoneToResults();
            bool groupByWorkshopSelected = resultsPage.GetGroupByPage(GROUP_BY_WORKSHOP);
            Assert.IsTrue(rawMatByGroupSelected && groupByWorkshopSelected, String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by group-- Workshop FaF"));

            // Group by customer
            filterAndFavorite = resultsPage.Back();
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.RawMaterialByGroup, true);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.GroupBy, GROUP_BY_CUSTOMER);

            resultsPage = filterAndFavorite.DoneToResults();
            bool groupByCustomerSelected = resultsPage.GetGroupByPage(GROUP_BY_CUSTOMER);
            Assert.IsTrue(rawMatByGroupSelected && groupByCustomerSelected, String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by group-- Customer FaF"));

            // Group by recipe type
            filterAndFavorite = resultsPage.Back();
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.RawMaterialByGroup, true);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.GroupBy, GROUP_BY_RECIPE_TYPE);

            resultsPage = filterAndFavorite.DoneToResults();
            bool groupByRecipeTypeSelected = resultsPage.GetGroupByPage(GROUP_BY_RECIPE_TYPE);
            Assert.IsTrue(rawMatByGroupSelected && groupByRecipeTypeSelected, String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by group-- Recipe type FaF"));
        }
        [Timeout(_timeout)]
        [TestMethod]

        public void PU_NEEDS_RawMatByWorkhop()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();

            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string customerName3 = TestContext.Properties["Needs_Customer3"].ToString();

            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();

            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = TestContext.Properties["Needs_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            if (!isServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);
            }

            if (!isServiceValide)
            {
                CreateFlight(homePage, siteACE, flightNameValide, customerName3, guestType2, serviceNameFlightValide, true);
            }

            if (!isServiceNormal || !isServiceValide)
            {
                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            }

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();

            // On vérifie que les couples subgroup/recette de la page sont corrects
            Assert.IsTrue(resultPage.GetSubgroupAndRecipeAssociated(), "Les informations affichées dans la vue 'Raw mat by workshop' ne sont pas correctes.");
        }
        [Timeout(_timeout)]
        [TestMethod]

        public void PU_NEEDS_RawMatByWorkshop_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "Raw mat by workshop" et par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.RawMaterialByWorkshop, true);

            var resultsPage = filterAndFavorite.DoneToResults();
            bool rowMatByWorkshopSelected = resultsPage.GetRawMatByPage(RAW_MAT_WORKSHOP);

            // Assert
            Assert.IsTrue(rowMatByWorkshopSelected, String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by workshop FaF"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_RawMatBySupplier()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();

            //Arrange

            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            if (!isServiceNormal || !isServiceValide)
            {
                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            }

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();

            // On récupère les couples groupe/items de la page
            Dictionary<string, List<string>> mapSupplierItems = resultPage.GetGroupAndItemsAssociated();

            var itemPage = homePage.GoToPurchasing_ItemPage();

            foreach (var entry in mapSupplierItems)
            {
                foreach (string itemName in entry.Value)
                {
                    if (!String.IsNullOrEmpty(itemName))
                    {
                        itemPage.ResetFilter();
                        itemPage.Filter(ItemPage.FilterType.Search, itemName);
                        var itemGeneralInformationPage = itemPage.ClickOnFirstItem();

                        Assert.AreEqual(entry.Key, itemGeneralInformationPage.GetPackagingSupplierBySite(siteACE), "Le supplier " + entry.Key + " associé à l'item " + itemName + " dans la page Raw mat by supplier n'est pas correct.");
                        itemPage = itemGeneralInformationPage.BackToList();
                    }
                }
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_RawMatBySupplier_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "Raw mat by supplier" et par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.RawMaterialBySupplier, true);

            var resultsPage = filterAndFavorite.DoneToResults();
            bool rowMatBySupplierSelected = resultsPage.GetRawMatByPage(RAW_MAT_SUPPLIER);

            // Assert
            Assert.IsTrue(rowMatBySupplierSelected, String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by supplier FaF"));
        }
        [Timeout(_timeout)]
        [TestMethod]


        public void PU_NEEDS_RawMatByRecipe()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();

            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string customerName3 = TestContext.Properties["Needs_Customer3"].ToString();

            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Needs_ServiceFlightValide"].ToString();

            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = TestContext.Properties["Needs_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            if (!isServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);
            }

            if (!isServiceValide)
            {
                CreateFlight(homePage, siteACE, flightNameValide, customerName3, guestType2, serviceNameFlightValide, true);
            }

            if (!isServiceNormal || !isServiceValide)
            {
                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            }

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();

            // Récupération de l'association des recettes et de leurs items
            var mapRecipeItems = resultPage.GetRecipeAndItemsAssociated();

            var recipesPage = homePage.GoToMenus_Recipes();

            foreach (var entry in mapRecipeItems)
            {
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, entry.Key);
                var recipeGeneralInfoPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfoPage.SelectVariantWithSite(siteACE);

                List<string> itemList = recipeVariantPage.GetIngredients();

                bool result = itemList.All(s => entry.Value.Contains(s)) && entry.Value.All(s => itemList.Contains(s));

                Assert.AreEqual(entry.Value.Count, itemList.Count, "Le nombre d'ingrédients de la recette " + entry.Key + " n'est pas égal à celui issu de la vue 'Raw mat by recipe'");
                Assert.IsTrue(result, "Les ingrédients de la recette " + entry.Key + " que sont pas les mêmes que ceux de la vue 'Raw mat by recipe'");

                recipesPage = recipeVariantPage.BackToList();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_RawMatByRecipe_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "Raw mat by recipe type" et par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.RawMaterialByRecipe, true);

            var resultsPage = filterAndFavorite.DoneToResults();
            bool rowMatByRecipeSelected = resultsPage.GetRawMatByPage(RAW_MAT_RECIPE);

            // Assert
            Assert.IsTrue(rowMatByRecipeSelected, String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by recipe FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_RawMatByCustomer()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);

            Assert.IsTrue(isServiceNormal, "Le service " + serviceNameNormal + " n'est pas visible.");

            List<string> customerList = qtyAjustementPage.GetCustomerList();

            ResultPageNeeds resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_CUSTOMER);
            resultPage.UnfoldAll();

            List<string> listCustomersResults = resultPage.GetCustomerNames();

            bool result = customerList.All(s => listCustomersResults.Contains(s)) && listCustomersResults.All(s => customerList.Contains(s));

            //assert
            Assert.IsTrue(result, "Les customers affichés dans la page Raw mat by customer ne correspondent pas à ceux des services.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_RawMatByCustomer_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "Raw mat by recipe type" et par site
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.RawMaterialByCustomer, true);

            var resultsPage = filterAndFavorite.DoneToResults();
            bool rowMatByCustomerSelected = resultsPage.GetRawMatByPage(RAW_MAT_CUSTOMER);

            // Assert
            Assert.IsTrue(rowMatByCustomerSelected, String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by customer FaF"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_UnfoldAll()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();


            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            int count = resultPage.CountItemsRawMatGroup();
            Assert.AreNotEqual(0, count, "Il n'y a pas de groupes d'items visibles.");
            Assert.IsTrue(resultPage.IsFoldAll(), "Le détail des items visible avant d'avoir activé le UnfoldAll.");

            // On unfold le premier item (test arrow)
            resultPage.ShowDetail();
            Assert.IsTrue(resultPage.IsDetailVisible(), "L'option d'unfold via l'item n'est pas fonctionnelle.");

            WebDriver.Navigate().Refresh();
            qtyAjustementPage.PageSize("100");
            resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            resultPage.UnfoldAll();
            Assert.IsTrue(resultPage.IsUnfoldAll(), "L'option d'unfoldAll via le menu n'est pas fonctionnelle.");
        }

        [Timeout(_timeout)]
        [TestMethod]

        public void PU_NEEDS_FoldAll()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            Assert.AreNotEqual(0, resultPage.CountItemsRawMatGroup(), "Il n'y a pas de groupes d'items visibles.");
            Assert.IsTrue(resultPage.IsFoldAll(), "Le détail des items visible avant d'avoir activé le UnfoldAll.");

            // On unfold le premier item (test arrow)
            resultPage.ShowDetail();
            Assert.IsTrue(resultPage.IsDetailVisible(), "L'option d'unfold via l'item n'est pas fonctionnelle.");

            // On replie l'item
            resultPage.ShowDetail();
            Assert.IsFalse(resultPage.IsDetailVisible(), "L'option de fold via l'item n'est pas fonctionnelle.");

            // test avec menu
            WebDriver.Navigate().Refresh();
            qtyAjustementPage.PageSize("100");
            resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            resultPage.UnfoldAll();
            Assert.IsTrue(resultPage.IsUnfoldAll(), "L'option UnfoldAll via le menu n'est pas fonctionnelle.");

            resultPage.FoldAll();
            Assert.IsTrue(resultPage.IsFoldAll(), "L'option FoldAll via le menu n'est pas fonctionnelle.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_Show_UseCase()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Needs_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);

            if (!isServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);

                filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            }

            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            resultPage.UnfoldAll();
            var resultModalPage = resultPage.ShowUseCase();

            //assert
            Assert.IsTrue(resultModalPage.IsServiceVisible(), "La page des use cases n'est pas apparu.");
        }

        //Generate Supply order
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_NEEDS_GenerateSupplyOrder()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string location = TestContext.Properties["PlaceFrom"].ToString();

            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();

            //Arrange

            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            // On récupère la liste des items présents
            List<String> items = resultPage.GetItemNames();

            // Génération de l'outputForm
            string supplyOrderNumber = resultPage.GenerateSupplyOrder(location);
            var supplyOrderItemPage = resultPage.Submit();

            try
            {
                var supplyOrderItems = supplyOrderItemPage.GetItemNames();
                Assert.IsTrue(items.All(s => supplyOrderItems.Contains(s)), "Les items du supply order ne sont pas les mêmes que ceux du production management");
                var supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(SupplyOrderPage.FilterType.ByNumber, supplyOrderNumber);

                Assert.AreEqual(supplyOrderNumber, supplyOrderPage.GetFirstSONumber(), "Le supply order générée n'est pas dans la liste des supply order");
            }
            finally
            {
                supplyOrderItemPage.Close();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]

        public void PU_NEEDS_Filter_MakeFavoriteFoodPacket()
        {
            // Prepare
            string favoriteName = "FoodPacketTest";
            //Arrange
            var homePage = LogInAsAdmin();

            //Page Make Favorite et Avoir des données
            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            if (filterAndFavorite.IsFavoritePresent(favoriteName))
            {
                filterAndFavorite.DeleteFavorite(favoriteName);
            }

            //Appliquer le filtre Food Packets puis cliquer sur Make Favorite
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.FoodPacket, "None");
            string avant = filterAndFavorite.GetFoodPacket();
            Assert.AreEqual("1 of 1 Food Packets selected", avant);
            filterAndFavorite.MakeFavorite(favoriteName);

            //Le filtre est enregistré
            ItemPage items = filterAndFavorite.GoToPurchasing_ItemPage();
            filterAndFavorite = items.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            Assert.IsTrue(filterAndFavorite.IsFavoritePresent(favoriteName));
            filterAndFavorite.SelectFoodPacket(null);
            var apres = filterAndFavorite.GetFoodPacket();
            Assert.AreEqual("Nothing Selected", filterAndFavorite.GetFoodPacket());
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.FoodPacket, "None");
            Assert.AreEqual(avant, filterAndFavorite.GetFoodPacket(), "Normalement 1 of 1");
            filterAndFavorite.DeleteFavorite(favoriteName);
        }
        [Timeout(_timeout)]
        [TestMethod]

        public void PU_NEEDS_Details_FilterByCustomer_Results()
        {
            // prepare
            string customerName1 = TestContext.Properties["Needs_Customer1"].ToString();
            DateTime dateFrom = DateUtils.Now.AddDays(-10);
            // arrange
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.DateFrom, dateFrom);
            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            ResultPageNeeds resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterByCustomer();
            // filtre by customer
            filterAndFavorite.Filter(FilterAndFavoritesNeedsPage.FilterType.Customers, customerName1);
            string item = resultPage.GetNeedsResultName();
            bool isContaining = item.Contains(customerName1);
            // assert
            Assert.IsTrue(isContaining, "Customer name ne correspond au customer filtré ");
        }
        [Timeout(_timeout)]
        [TestMethod]


        public void PU_NEEDS_Details_FilterByRawMatByGroupAndCustomer_Results()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            ResultPageNeeds resultPage = qtyAjustementPage.GoToResultPage();

            // Get the initial list of items before applying the filter
            List<string> listItemGroup = resultPage.GetNeedsResultsContant();

            // Step 1: Select "Raw mat by group" from the filter dropdown
            var filterDropdownElement = resultPage.GetFilterDropdown();
            var filterDropdown = new SelectElement(filterDropdownElement);
            filterDropdown.SelectByText("Raw mat by group");

            // Step 2: Verify the list on the right is displayed
            Assert.IsTrue(resultPage.IsRightListDisplayed(), "The right list is not displayed after selecting 'Raw mat by group'.");

            // Step 3: Select "Customer" from the right dropdown list
            resultPage.FilterByCustomer();

            // Step 4: Retrieve the list of results after filtering by Customer
            List<string> listCustomersResults = resultPage.GetNeedsResultsContant();

            // Step 5: Verify that the result list has changed
            Assert.AreNotEqual(listCustomersResults, listItemGroup, "The list didn't change after applying the Customer filter.");

            // Step 6: (Optional) Verify that the results contain customer-related data
            Assert.IsTrue(listCustomersResults.All(result => result.Contains("Customer")),
                "Not all results are related to Customers after filtering.");
        }



        // __________________________________________ Utilitaire ________________________________________________
        public void CreateFlight(HomePage homePage, string site, string flightName, string customerName, string guestType, string serviceName, bool isValid = false)
        {
            // prepare
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightName);

            if (flightPage.CheckTotalNumber() == 0)
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightName, customerName, aircraft, site, siteBCN);

                //Edit the first flight
                var editPage = flightPage.EditFirstFlight(flightName);
                editPage.AddGuestType(guestType);
                editPage.AddService(serviceName);
                editPage.CloseViewDetails();

                if (isValid)
                {
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightName);
                    flightPage.SetNewState("V");
                }
            }
            else
            {
                var editPage = flightPage.EditFirstFlight(flightName);

                bool isFlightValid = false;

                if (isValid)
                {
                    isFlightValid = editPage.GetFlightStatus("V");
                }

                if (editPage.GetGuestType().Equals("") || !editPage.GetGuestType().Equals(guestType))
                {
                    editPage.AddGuestType(guestType);
                    editPage.AddService(serviceName);
                }

                if (editPage.GetServiceName(isFlightValid).Equals("") || !editPage.GetServiceName(isFlightValid).Equals(serviceName))
                {
                    editPage.AddService(serviceName);
                }

                editPage.CloseViewDetails();

                if (isValid && !isFlightValid)
                {
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightName);
                    flightPage.SetNewState("V");
                    WebDriver.Navigate().Refresh();
                }
            }

            var schedulePage = homePage.GoToFlights_FlightSelectionPage();
            schedulePage.ResetFilter();
            schedulePage.Filter(SchedulePage.FilterType.Site, site);

            schedulePage.UnfoldAll();
            schedulePage.SetFlightProduced(flightName, true);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_NEEDS_Filter_SortByService()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            string serviceNameVacuum = TestContext.Properties["Needs_ServiceVacuum"].ToString();
                        //Arrange
            HomePage homePage = LogInAsAdmin();
                        //Act
            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            bool isServiceNormalPresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(isServiceNormalPresent, $"Le service {serviceNameNormal} n'est pas visible.");
            bool isServiceVaccumPresent = qtyAjustementPage.IsServicePresent(serviceNameVacuum);
            Assert.IsTrue(isServiceVaccumPresent, $"Le service {serviceNameVacuum} n'est pas visible.");
            // Sort by service
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.SortBy, "Service");
            bool isSortedByService = qtyAjustementPage.IsSortedByService();
            //assert
            Assert.IsTrue(isSortedByService, string.Format(MessageErreur.FILTRE_ERRONE, "Sort by service"));
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_NEEDS_Filter_From_Result()
        {
            // verification de fonc: filtre sur l interface QA 
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Needs_ServiceNormal"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            FilterAndFavoritesNeedsPage filterAndFavorite = homePage.GoToPurchasing_NeedsPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsNeedsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsNeedsPage.FilterType.DateFrom, DateTime.Today.AddDays(-1));
            bool isServiceNormalPresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(isServiceNormalPresent, $"Le service {serviceNameNormal} n'est pas visible , verifier la creation du vol ou la fonctionnaliter de filtre from ");

        }

        #region "SQL INSERTION FOR FLIGHT PRECONDTION"
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

        private void InsertPrevalFlight(int flightDeliverieId, string date, string aircraft, string airoportFrom, string airoportTo)
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
        private void InsertValidFlight(int flightDeliverieId, string date, string aircraft, string airoportFrom, string airoportTo)
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
            1,
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
        private void InsertInvoiceFlight(int flightDeliverieId, string date, string aircraft, string airoportFrom, string airoportTo)
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
            1,
            1,
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
            //Insert FlightDetails For Flight
            string queryFlightDetails = @"
            DECLARE @serviceId INT;
            SELECT TOP 1 @serviceId = Id FROM Services WHERE Name LIKE @serviceName;

               Insert into [FlightDetails] (
              [FlightLegToGuestTypeId]
              ,[ServiceId]
              ,[Mode]
              ,[Value]
              ,[FinalQuantity]
              ,[ComputedQuantity]
              ,[CreationDate]
              ,[ModificationDate]
              ,[ProductionStatus]
              ,[IsManual]
              ,[IsSpml])
	          values(
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
 SELECT SCOPE_IDENTITY();
                "
            ;
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

        private void DeleteFlightDeliverie(int flightdeliveryId)
        {
            string query = @"
        DELETE FROM FlightDeliveries 
        WHERE Id = @flightdeliveryId;";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("flightdeliveryId", flightdeliveryId));
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

        private void DeleteFlightHistorie(int flightId)
        {
            string query = @"
            delete from FlightHistories where FlightId = @flightId;";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("flightId", flightId));
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

        private void DeleteFlightProperties(List<int> flightsIds)
        {
            foreach (var flightId in flightsIds)
            {
                DeleteFlightPropertie(flightId);
            }
        }

        private void DeleteFlightHistories(List<int> flightsIds)
        {
            foreach (var flightId in flightsIds)
            {
                DeleteFlightHistorie(flightId);
            }
        }

        /// </summary>

    }
    #endregion
}
