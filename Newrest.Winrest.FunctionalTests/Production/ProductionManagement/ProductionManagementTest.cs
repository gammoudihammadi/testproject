using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Schedule;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.User;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Page = UglyToad.PdfPig.Content.Page;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;

namespace Newrest.Winrest.FunctionalTests.Production
{
    [TestClass]
    public class ProductionManagementTest : TestBase
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
        private static Random random = new Random();

        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
        /// 
        [TestInitialize]
        public override void TestInitialize()
        {

            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(PR_PRODMAN_Filter_RecipeType):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_Filter_BatchStartTime):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_BatchEndTime):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_ShowValidatedFlightsOnly):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_Valid_J1();
                    break;
                case nameof(PR_PRODMAN_Filter_Workshops):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_Filter_GuestType):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_Filter_Site):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Normal_MAD_J1();
                    break;
                case nameof(PR_PRODMAN_MakeFavorite):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    break;
                case nameof(PR_PRODMAN_Filter_ItemGroup):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    break;
                case nameof(PR_PRODMAN_Filter_Service):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_Filter_Customers):
                    //Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_Filter_ItemGroup_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Edit_Qty_Adjustement):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_UnfoldAll):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    break;
                case nameof(PR_PRODMAN_RawMatByRecipe):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_RawMatByCustomer):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1();
                    break;
                case nameof(PR_PRODMAN_FoldAll):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    break;
                case nameof(PR_PRODMAN_Show_UseCase):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    break;
                case nameof(PR_PRODMAN_GenerateOutputForm):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    break;
                case nameof(PR_PRODMAN_Print_RawMaterialByWorkshop_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_RecipeReportV2_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_RawMatByGroup):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_RawMatBySupplier):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1();
                    break;
                case nameof(PR_PRODMAN_RazAllQuantities):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_ResetAllQuantities):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_Calcul_OF_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Vaccum();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1();
                    break;
                case nameof(PR_PRODMAN_Filter_ServiceCategorie):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_RawMatByWorkhop):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_Print_HACCPPlatingByGroup_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_HACCPPlating_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_HACCPHotKitchen_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    break;
                case nameof(PR_PRODMAN_Print_DatasheetByGuest_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_Assembly_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_NoCancelOrInvoiceFlight):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_HACCPSanitization_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_HACCPSlice_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_HACCPThawing_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_HACCPTraySetup_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_RawMaterialByGroup_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_RecipeReportDetailedV2_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_FD_NotRounded):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    break;
                case nameof(PR_PRODMAN_Filter_ShowVacuumProductionOnly):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    break;
                case nameof(PR_PRODMAN_PrintCheck_DatasheetByguest_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_PrevalidateOrValidateFlightOnly):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1(); 
                    break;
                case nameof(PR_PRODMAN_PrintCheck_DatasheetByRecipe_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_ItemHide_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_Filter_Customers_FaF):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;

                case nameof(PR_PRODMAN_RawMatByWorkshop_FaF):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_ServiceGlobalView3_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_DatasheetByRecipe_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_PrintCheck_AssemblyReportServiceName_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Print_HACCPPlatingByRecipe_NewVersion):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_FromTo):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight_Before2Days();
                    break;
                case nameof(PR_PRODMAN_Filter_ItemSubgroup_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1();
                    break;
                case nameof(PR_PRODMAN_GenerateSupplyOrder):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Quantity):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight();
                    break;
                case nameof(PR_PRODMAN_PrintCheck_RawMatByWorkshop_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_Customer_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_ItemSubgroup_wSubgroup_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_Service_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_GuestType_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_RecipeType_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_ShowVaccumOnly_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_ShowNormalOnly_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                    break;
                case nameof(PR_PRODMAN_OpenService_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_UseCaseCRecette_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_PrintCheck_RecipeReportdetailedV2_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_PrintCheck_HaccpSanitization_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_ServiceGlobalView1_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_ShowNormalAndVacuumProduction):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    break;
                case nameof( PR_PRODMAN_ServiceGlobalView4_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_PrintCheck_HaccpPlatingGroupByRecipe_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Calcul_SO_Result_Storage):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Calcul_SO_Result_Total_wo_VAT):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_ShowNormalProductionOnly):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    break;
                case nameof(PR_PRODMAN_PrintCheck_HaccpThawing_Result):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_ServiceType_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1();
                    break;
                case nameof(PR_PRODMAN_CheckResultPage_QA):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof(PR_PRODMAN_Filter_Service_FaF):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
                    break;
                case nameof( PR_PRODMAN_Filter_ServiceCategorie_FaF):
                    Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1();
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
                case nameof(PR_PRODMAN_Filter_RecipeType):
                case nameof(PR_PRODMAN_Filter_BatchStartTime):
                case nameof(PR_PRODMAN_Filter_BatchEndTime):
                case nameof(PR_PRODMAN_Filter_ShowValidatedFlightsOnly):
                case nameof(PR_PRODMAN_Filter_Workshops):
                case nameof(PR_PRODMAN_Filter_GuestType):
                case nameof(PR_PRODMAN_Filter_Site):
                case nameof(PR_PRODMAN_MakeFavorite):
                case nameof(PR_PRODMAN_Filter_ItemGroup):
                case nameof(PR_PRODMAN_Filter_Service):
                case nameof(PR_PRODMAN_Filter_Customers):
                case nameof(PR_PRODMAN_Filter_ItemGroup_QA):
                case nameof(PR_PRODMAN_Edit_Qty_Adjustement):
                case nameof(PR_PRODMAN_UnfoldAll):
                case nameof(PR_PRODMAN_RawMatByRecipe):
                case nameof(PR_PRODMAN_FoldAll):
                case nameof(PR_PRODMAN_RawMatByCustomer):
                case nameof(PR_PRODMAN_Show_UseCase):
                case nameof(PR_PRODMAN_GenerateOutputForm):
                case nameof(PR_PRODMAN_Print_RawMaterialByWorkshop_NewVersion):
                case nameof(PR_PRODMAN_Print_RecipeReportV2_NewVersion):
                case nameof(PR_PRODMAN_RawMatByGroup):
                case nameof(PR_PRODMAN_RawMatBySupplier):
                case nameof(PR_PRODMAN_ResetAllQuantities):
                case nameof(PR_PRODMAN_Calcul_OF_Result):
                case nameof(PR_PRODMAN_RazAllQuantities):
                case nameof(PR_PRODMAN_Filter_ServiceCategorie):
                case nameof(PR_PRODMAN_RawMatByWorkhop):
                case nameof(PR_PRODMAN_Print_HACCPPlatingByGroup_NewVersion):
                case nameof(PR_PRODMAN_Print_HACCPPlating_NewVersion):
                case nameof(PR_PRODMAN_Print_HACCPHotKitchen_NewVersion):
                case nameof(PR_PRODMAN_Print_DatasheetByGuest_NewVersion):
                case nameof(PR_PRODMAN_Print_Assembly_NewVersion):
                case nameof(PR_PRODMAN_NoCancelOrInvoiceFlight):
                case nameof(PR_PRODMAN_Print_HACCPSanitization_NewVersion):
                case nameof(PR_PRODMAN_Print_HACCPSlice_NewVersion):
                case nameof(PR_PRODMAN_Print_HACCPThawing_NewVersion):
                case nameof(PR_PRODMAN_Print_HACCPTraySetup_NewVersion):
                case nameof(PR_PRODMAN_Print_RawMaterialByGroup_NewVersion):
                case nameof(PR_PRODMAN_Print_RecipeReportDetailedV2_NewVersion):
                case nameof(PR_PRODMAN_FD_NotRounded):
                case nameof(PR_PRODMAN_PrintCheck_DatasheetByguest_Result):
                case nameof(PR_PRODMAN_PrevalidateOrValidateFlightOnly):
                case nameof(PR_PRODMAN_Filter_ItemHide_Result):
                case nameof(PR_PRODMAN_Filter_Customers_FaF):
                case nameof(PR_PRODMAN_RawMatByWorkshop_FaF):
                case nameof(PR_PRODMAN_ServiceGlobalView3_QA):
                case nameof(PR_PRODMAN_Print_DatasheetByRecipe_NewVersion):
                case nameof(PR_PRODMAN_PrintCheck_AssemblyReportServiceName_Result):
                case nameof(PR_PRODMAN_Print_HACCPPlatingByRecipe_NewVersion):
                case nameof(PR_PRODMAN_Filter_FromTo):
                case nameof(PR_PRODMAN_Filter_ItemSubgroup_Result):
                case nameof(PR_PRODMAN_GenerateSupplyOrder):
                case nameof(PR_PRODMAN_Quantity):
                case nameof(PR_PRODMAN_PrintCheck_DatasheetByRecipe_Result):
                case nameof(PR_PRODMAN_PrintCheck_RawMatByWorkshop_Result):
                case nameof(PR_PRODMAN_Filter_Customer_QA):
                case nameof(PR_PRODMAN_Filter_ItemSubgroup_wSubgroup_QA):
                case nameof(PR_PRODMAN_Filter_Service_QA):
                case nameof(PR_PRODMAN_Filter_GuestType_QA):
                case nameof(PR_PRODMAN_Filter_RecipeType_QA):
                case nameof(PR_PRODMAN_Filter_ShowVaccumOnly_QA):
                case nameof(PR_PRODMAN_Filter_ShowNormalOnly_QA):
                case nameof(PR_PRODMAN_OpenService_QA):
                case nameof(PR_PRODMAN_Filter_UseCaseCRecette_Result):
                case nameof(PR_PRODMAN_PrintCheck_RecipeReportdetailedV2_Result):
                case nameof(PR_PRODMAN_PrintCheck_HaccpSanitization_Result):
                case nameof(PR_PRODMAN_ServiceGlobalView1_QA):
                case nameof(PR_PRODMAN_Filter_ShowNormalAndVacuumProduction):
                case nameof(PR_PRODMAN_Filter_ShowVacuumProductionOnly):
                case nameof(PR_PRODMAN_ServiceGlobalView4_QA):
                case nameof(PR_PRODMAN_PrintCheck_HaccpPlatingGroupByRecipe_Result):
                case nameof(PR_PRODMAN_Calcul_SO_Result_Storage):
                case nameof(PR_PRODMAN_Calcul_SO_Result_Total_wo_VAT):
                case nameof(PR_PRODMAN_Filter_ShowNormalProductionOnly):
                case nameof(PR_PRODMAN_PrintCheck_HaccpThawing_Result):
                case nameof(PR_PRODMAN_Filter_ServiceType_QA):
                case nameof(PR_PRODMAN_CheckResultPage_QA):
                case nameof(PR_PRODMAN_Filter_Service_FaF):
                case nameof(PR_PRODMAN_Filter_ServiceCategorie_FaF):
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


        // -------------------- Test Init --------------------------------------------------------------- 

        public void Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = "" + TestContext.Properties["Prodman_FlightNormal"].ToString() + DateTime.Today.ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);

            TestContext.Properties[string.Format("FlightDeliveriesId1")] = InsertFlightDeliveries(siteACE, customerName1, flightNameNormal);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId1")], DateTime.Today.ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId1")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId1")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId1")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId1")]);
            TestContext.Properties[string.Format("FlightDetailsId1")] = InsertFlightDetailsWithService(serviceNameNormal, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId1")]);

        }
        public void Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = "" + TestContext.Properties["Prodman_FlightNormal"].ToString() + DateTime.Today.AddDays(1).ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);

            TestContext.Properties[string.Format("FlightDeliveriesId2")] = InsertFlightDeliveries(siteACE, customerName1, flightNameNormal);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId2")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId2")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId2")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId2")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId2")]);
            TestContext.Properties[string.Format("FlightDetailsId2")] = InsertFlightDetailsWithService(serviceNameNormal, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId2")]);

        }
        public void Test_Initialize_PR_PRODMAN_CreateFlight_normal_flightJ1_INVOICED()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = "" + TestContext.Properties["Prodman_FlightNormal"].ToString() + DateTime.Today.AddDays(1).ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);

            TestContext.Properties[string.Format("FlightDeliveriesId2")] = InsertFlightDeliveries(siteACE, customerName1, flightNameNormal);
            InsertInvoiceFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId2")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId2")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId2")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId2")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId2")]);
            TestContext.Properties[string.Format("FlightDetailsId2")] = InsertFlightDetailsWithService(serviceNameNormal, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId2")]);

        }
        public void Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight_Valid()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = "" + TestContext.Properties["Prodman_FlightNormal"].ToString() + DateTime.Today.ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);

            TestContext.Properties[string.Format("FlightDeliveriesId3")] = InsertFlightDeliveries(siteACE, customerName1, flightNameNormal);
            InsertValidFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId3")], DateTime.Today.ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId3")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId3")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId3")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId3")]);
            TestContext.Properties[string.Format("FlightDetailsId3")] = InsertFlightDetailsWithService(serviceNameNormal, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId3")]);

        }
  

        public void Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType2()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            string flightNameValide = TestContext.Properties["Prodman_FlightValide"].ToString() + DateTime.Today.ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);

            TestContext.Properties[string.Format("FlightDeliveriesId4")] = InsertFlightDeliveries(siteACE, customerName3, flightNameValide);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId4")], DateTime.Today.ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId4")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId4")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId4")] = InsertFlightLegToGuestTypes(guestType2, (int)TestContext.Properties[string.Format("FlightLegsId4")]);
            TestContext.Properties[string.Format("FlightDetailsId4")] = InsertFlightDetailsWithService(serviceNameFlightValide, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId4")]);

        }

        public void Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            string flightNameValide = TestContext.Properties["Prodman_FlightValide"].ToString() + DateTime.Today.ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);

            TestContext.Properties[string.Format("FlightDeliveriesId5")] = InsertFlightDeliveries(siteACE, customerName3, flightNameValide);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId5")], DateTime.Today.ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId5")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId5")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId5")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId5")]);
            TestContext.Properties[string.Format("FlightDetailsId5")] = InsertFlightDetailsWithService(serviceNameFlightValide, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId5")]);

        }
        public void Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_J1()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            string flightNameValide = TestContext.Properties["Prodman_FlightValide"].ToString() + DateTime.Today.AddDays(1).ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);

            TestContext.Properties[string.Format("FlightDeliveriesId6")] = InsertFlightDeliveries(siteACE, customerName3, flightNameValide);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId6")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId6")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId6")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId6")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId6")]);
            TestContext.Properties[string.Format("FlightDetailsId6")] = InsertFlightDetailsWithService(serviceNameFlightValide, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId6")]);

        }
        public void Test_Initialize_PR_PRODMAN_CreateFlight_Validated_flight_guestType1_Valid_J1()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            string flightNameValide = TestContext.Properties["Prodman_FlightValide"].ToString() + DateTime.Today.AddDays(1).ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);

            TestContext.Properties[string.Format("FlightDeliveriesId7")] = InsertFlightDeliveries(siteACE, customerName3, flightNameValide);
            InsertValidFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId7")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId7")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId7")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId7")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId7")]);
            TestContext.Properties[string.Format("FlightDetailsId7")] = InsertFlightDetailsWithService(serviceNameFlightValide, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId7")]);

        }

        public void Test_Initialize_PR_PRODMAN_CreateFlight_Normal_MAD_J1()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormalMAD = TestContext.Properties["Prodman_ServiceNormalMAD"].ToString();
            string flightNameNormalMAD = "" + TestContext.Properties["Prodman_FlightNormalMAD"].ToString() + DateTime.Today.AddDays(1).ToString("yyyy-MM-dd") + "_" + random.Next(0, 1000);


            TestContext.Properties[string.Format("FlightDeliveriesId8")] = InsertFlightDeliveries(siteMAD, customerName1, flightNameNormalMAD);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId8")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteMAD, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId8")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId8")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId8")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId8")]);
            TestContext.Properties[string.Format("FlightDetailsId8")] = InsertFlightDetailsWithService(serviceNameNormalMAD, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId8")]);

        }
        public void Test_Initialize_PR_PRODMAN_CreateFlight_normal_flight_Before2Days()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
          
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateTime.Today.AddDays(-2).ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);

            TestContext.Properties[string.Format("FlightDeliveriesId12")] = InsertFlightDeliveries(siteACE, customerName1, flightNameNormal);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId12")], DateTime.Today.AddDays(-2).ToString("yyyy/MM/dd"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId12")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId12")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId12")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId12")]);
            TestContext.Properties[string.Format("FlightDetailsId12")] = InsertFlightDetailsWithService(serviceNameNormal, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId12")]);
        }
        public void Test_Initialize_PR_PRODMAN_CreateFlight_Normal_MAD()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormalMAD = TestContext.Properties["Prodman_ServiceNormalMAD"].ToString();
            string flightNameNormalMAD = "" + TestContext.Properties["Prodman_FlightNormalMAD"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);
            string serviceMAD;

            TestContext.Properties[string.Format("FlightDeliveriesId9")] = InsertFlightDeliveries(siteMAD, customerName1, flightNameNormalMAD);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId9")], DateTime.Today.ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteMAD, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId9")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId9")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId9")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId9")]);
            TestContext.Properties[string.Format("FlightDetailsId9")] = InsertFlightDetailsWithService(serviceNameNormalMAD, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId9")]);

        }
        public void Test_Initialize_PR_PRODMAN_CreateFlight_Vaccum()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string serviceNameVacuum = TestContext.Properties["Prodman_ServiceVacuum"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();

            string customerName2 = TestContext.Properties["Prodman_Customer2"].ToString();

            string flightNameVacuum = TestContext.Properties["Prodman_FlightVacuum"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);

            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();

            TestContext.Properties[string.Format("FlightDeliveriesId10")] = InsertFlightDeliveries(siteACE, customerName2, flightNameVacuum);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId10")], DateTime.Today.ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId10")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId10")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId10")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId10")]);
            TestContext.Properties[string.Format("FlightDetailsId10")] = InsertFlightDetailsWithService(serviceNameVacuum, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId10")]);

        }
        public void Test_Initialize_PR_PRODMAN_CreateFlight_VaccumJ1()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string serviceNameVacuum = TestContext.Properties["Prodman_ServiceVacuum"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();

            string customerName2 = TestContext.Properties["Prodman_Customer2"].ToString();

            string flightNameVacuum = TestContext.Properties["Prodman_FlightVacuum"].ToString() + DateTime.Today.AddDays(1).ToString("dd/MM/yyyy") + "_" + random.Next(0, 1000);

            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string siteBCN = TestContext.Properties["SiteToFlight"].ToString();

            TestContext.Properties[string.Format("FlightDeliveriesId11")] = InsertFlightDeliveries(siteACE, customerName2, flightNameVacuum);
            InsertPrevalFlight((int)TestContext.Properties[string.Format("FlightDeliveriesId11")], DateTime.Today.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss"), aircraft, siteACE, siteBCN);
            TestContext.Properties[string.Format("FlightLegsId11")] = InsertFlightLegs((int)TestContext.Properties[string.Format("FlightDeliveriesId11")]);
            TestContext.Properties[string.Format("FlightLegToGuestTypesId11")] = InsertFlightLegToGuestTypes(guestType1, (int)TestContext.Properties[string.Format("FlightLegsId11")]);
            TestContext.Properties[string.Format("FlightDetailsId11")] = InsertFlightDetailsWithService(serviceNameVacuum, (int)TestContext.Properties[string.Format("FlightLegToGuestTypesId11")]);

        }


        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void PR_PRODMAN_SetConfigWinrest4_0()
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
            ClearCache();

            //Check all documents in application settings
            var applicationSettings = homePage.GoToApplicationSettings();
            applicationSettings.CheckAllDocumentsDisplay();

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

        [TestMethod]
        [Priority(1)]
        [Timeout(TestTimeout.Infinite)]
        public void PR_PRODMAN_CreateItems()
        {
            //Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();

            string itemName1 = TestContext.Properties["Prodman_Item1"].ToString();
            string itemName2 = TestContext.Properties["Prodman_Item2"].ToString();
            string itemName3 = TestContext.Properties["Prodman_Item3"].ToString();
            string itemName4 = TestContext.Properties["Prodman_Item4"].ToString();

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
            HomePage homePage = LogInAsAdmin();

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

        [TestMethod]
        [Priority(2)]
        [Timeout(TestTimeout.Infinite)]
        public void PR_PRODMAN_CreateRecipe()
        {
            //Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string recipeVariant = TestContext.Properties["Prodman_Needs_RecipeVariant"].ToString();

            string recipeName1 = TestContext.Properties["Prodman_RecipeName1"].ToString();
            string recipeType1 = TestContext.Properties["Prodman_Needs_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Prodman_Needs_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Prodman_Needs_Workshop1"].ToString();

            string recipeName2 = TestContext.Properties["Prodman_RecipeName2"].ToString();
            string recipeType2 = TestContext.Properties["Prodman_Needs_RecipeType2"].ToString();
            string recipeCookingMode2 = TestContext.Properties["Prodman_Needs_CookingMode2"].ToString();
            string recipeWorkshop2 = TestContext.Properties["Prodman_Needs_Workshop2"].ToString();

            string recipeName3 = TestContext.Properties["Prodman_RecipeName3"].ToString();
            string recipeType3 = TestContext.Properties["Prodman_Needs_RecipeType3"].ToString();
            string recipeWorkshop3 = TestContext.Properties["Prodman_Needs_Workshop3"].ToString();

            string itemName1 = TestContext.Properties["Prodman_Item1"].ToString();
            string itemName2 = TestContext.Properties["Prodman_Item2"].ToString();
            string itemName3 = TestContext.Properties["Prodman_Item3"].ToString();
            string itemName4 = TestContext.Properties["Prodman_Item4"].ToString();

            Random rnd = new Random();

            //Arrange           
            //Arrange
            HomePage homePage = LogInAsAdmin();

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
        [Timeout(TestTimeout.Infinite)]
        public void PR_PRODMAN_CreateDatasheet()
        {
            //Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string datasheetGuestType = TestContext.Properties["DatasheetGuestType"].ToString();

            string recipeName1 = TestContext.Properties["Prodman_RecipeName1"].ToString();
            string recipeName2 = TestContext.Properties["Prodman_RecipeName2"].ToString();
            string recipeName3 = TestContext.Properties["Prodman_RecipeName3"].ToString();

            string datasheetName1 = TestContext.Properties["Prodman_DatasheetName1"].ToString();
            string datasheetNameMAD = TestContext.Properties["Prodman_DatasheetNameMAD"].ToString();
            string datasheetName2 = TestContext.Properties["Prodman_DatasheetName2"].ToString();
            string datasheetName3 = TestContext.Properties["Prodman_DatasheetName3"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();

            // ---------------------------------------------- datasheet 1 --------------------------------------------------
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName1);
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);

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
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            string datasheet1 = datasheetPage.GetFirstDatasheetName();

            // ---------------------------------------------- datasheet MAD --------------------------------------------------

            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetNameMAD);
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);

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
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            string datasheetMAD = datasheetPage.GetFirstDatasheetName();

            // ---------------------------------------------- datasheet 2 --------------------------------------------------

            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName2);
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);

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
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            string datasheet2 = datasheetPage.GetFirstDatasheetName();

            // ---------------------------------------------- datasheet 3 --------------------------------------------------

            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName3);
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);

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
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            string datasheet3 = datasheetPage.GetFirstDatasheetName();

            // -------------------------------------------------- Assert -------------------------------------------------
            Assert.AreEqual(datasheetName1, datasheet1, "La datasheet " + datasheetName1 + " n'existe pas.");
            Assert.AreEqual(datasheetNameMAD, datasheetMAD, "La datasheet " + datasheetNameMAD + " n'existe pas.");
            Assert.AreEqual(datasheetName2, datasheet2, "La datasheet " + datasheetName2 + " n'existe pas.");
            Assert.AreEqual(datasheetName3, datasheet3, "La datasheet " + datasheetName3 + " n'existe pas.");
        }

        [TestMethod]
        [Priority(4)]
        [Timeout(TestTimeout.Infinite)]
        public void PR_PRODMAN_CreateCustomer()
        {
            // Prepare
            string customerType = TestContext.Properties["CustomerType"].ToString();

            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string customerCode1 = TestContext.Properties["Prodman_CustomerCode1"].ToString();

            string customerName2 = TestContext.Properties["Prodman_Customer2"].ToString();
            string customerCode2 = TestContext.Properties["Prodman_CustomerCode2"].ToString();

            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();
            string customerCode3 = TestContext.Properties["Prodman_CustomerCode3"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
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
        [Timeout(TestTimeout.Infinite)]
        public void PR_PRODMAN_CreateService()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();

            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string customerName2 = TestContext.Properties["Prodman_Customer2"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();

            string datasheetName1 = TestContext.Properties["Prodman_DatasheetName1"].ToString();
            string datasheetNameMAD = TestContext.Properties["Prodman_DatasheetNameMAD"].ToString();
            string datasheetName2 = TestContext.Properties["Prodman_DatasheetName2"].ToString();
            string datasheetName3 = TestContext.Properties["Prodman_DatasheetName3"].ToString();

            string serviceCategorie1 = TestContext.Properties["Prodman_Needs_ServiceCategory1"].ToString();
            string serviceCategorie2 = TestContext.Properties["Prodman_Needs_ServiceCategory2"].ToString();
            string serviceCategorie3 = TestContext.Properties["Prodman_Needs_ServiceCategory3"].ToString();

            string serviceType = TestContext.Properties["Prodman_Needs_ServiceType"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameNormalMAD = TestContext.Properties["Prodman_ServiceNormalMAD"].ToString();
            string serviceNameVacuum = TestContext.Properties["Prodman_ServiceVacuum"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();
            string serviceNameSansDatasheet = TestContext.Properties["Prodman_ServiceSansDatasheet"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            // -------------------------------------------- service normal --------------------------------------------------
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameNormal);
            servicePage.Filter(ServicePage.FilterType.ShowOnlyActive, true);
            int counterService1 = servicePage.CheckTotalNumber();
            if (counterService1 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersUncheck, true);
                servicePage.Filter(ServicePage.FilterType.SitesUncheck, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesUncheck, true);
            }
            counterService1 = servicePage.CheckTotalNumber();
            if (counterService1 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.SitesCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesCheckAll, true);
            }
            counterService1 = servicePage.CheckTotalNumber();
            if (counterService1 == 0)
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
                servicePricePage.SearchPriceForCustomer(customerName1, siteACE, DateUtils.Now.AddDays(-11), DateUtils.Now.AddMonths(2), "20", datasheetName1);

                var serviceGeneralInformationsPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameNormal);
            int counterService2 = servicePage.CheckTotalNumber();
            if (counterService2 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersUncheck, true);
                servicePage.Filter(ServicePage.FilterType.SitesUncheck, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesUncheck, true);
            }
            counterService2 = servicePage.CheckTotalNumber();
            if (counterService2 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.SitesCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesCheckAll, true);
            }
            string serviceNormal = servicePage.GetFirstServiceName();

            // -------------------------------------------- service normal MAD --------------------------------------------------

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameNormalMAD);
            int counterService3 = servicePage.CheckTotalNumber();
            if (counterService3 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersUncheck, true);
                servicePage.Filter(ServicePage.FilterType.SitesUncheck, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesUncheck, true);
            }
            counterService3 = servicePage.CheckTotalNumber();
            if (counterService3 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.SitesCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesCheckAll, true);
            }
            counterService3 = servicePage.CheckTotalNumber();
            if (counterService3 == 0)
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
            int counterService4 = servicePage.CheckTotalNumber();
            if (counterService4 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersUncheck, true);
                servicePage.Filter(ServicePage.FilterType.SitesUncheck, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesUncheck, true);
            }
            counterService4 = servicePage.CheckTotalNumber();
            if (counterService4 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.SitesCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesCheckAll, true);
            }
            string serviceNormalMAD = servicePage.GetFirstServiceName();

            // -------------------------------------------- service vacuum --------------------------------------------------
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameVacuum);
            int counterService5 = servicePage.CheckTotalNumber();
            if (counterService5 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersUncheck, true);
                servicePage.Filter(ServicePage.FilterType.SitesUncheck, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesUncheck, true);
            }
            counterService5 = servicePage.CheckTotalNumber();
            if (counterService5 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.SitesCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesCheckAll, true);
            }
            counterService5 = servicePage.CheckTotalNumber();
            if (counterService5 == 0)
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
            int counterService6 = servicePage.CheckTotalNumber();
            if (counterService6 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersUncheck, true);
                servicePage.Filter(ServicePage.FilterType.SitesUncheck, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesUncheck, true);
            }
            counterService6 = servicePage.CheckTotalNumber();
            if (counterService6 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.SitesCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesCheckAll, true);
            }
            string serviceVacuum = servicePage.GetFirstServiceName();

            // -------------------------------------------- service flight valide --------------------------------------------------

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameFlightValide);
            int counterService7 = servicePage.CheckTotalNumber();
            if (counterService7 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersUncheck, true);
                servicePage.Filter(ServicePage.FilterType.SitesUncheck, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesUncheck, true);
            }
            counterService7 = servicePage.CheckTotalNumber();
            if (counterService7 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.SitesCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesCheckAll, true);
            }
            counterService7 = servicePage.CheckTotalNumber();
            if (counterService7 == 0)
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
            int counterService8 = servicePage.CheckTotalNumber();
            if (counterService8 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersUncheck, true);
                servicePage.Filter(ServicePage.FilterType.SitesUncheck, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesUncheck, true);
            }
            counterService8 = servicePage.CheckTotalNumber();
            if (counterService8 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.SitesCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesCheckAll, true);
            }
            string serviceFlightValide = servicePage.GetFirstServiceName();

            // -------------------------------------------- service sans Datasheet --------------------------------------------------
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameSansDatasheet);
            int counterService9 = servicePage.CheckTotalNumber();
            if (counterService9 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersUncheck, true);
                servicePage.Filter(ServicePage.FilterType.SitesUncheck, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesUncheck, true);
            }
            counterService9 = servicePage.CheckTotalNumber();
            if (counterService9 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.SitesCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesCheckAll, true);
            }
            counterService9 = servicePage.CheckTotalNumber();
            if (counterService9 == 0)
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
            int counterService10 = servicePage.CheckTotalNumber();
            if (counterService10 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersUncheck, true);
                servicePage.Filter(ServicePage.FilterType.SitesUncheck, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesUncheck, true);
            }
            counterService10 = servicePage.CheckTotalNumber();
            if (counterService10 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.SitesCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesCheckAll, true);
            }
            string serviceSansDatasheet = servicePage.GetFirstServiceName();

            // -------------------------------------------------- Assert -------------------------------------------------
            Assert.IsTrue(serviceNormal.Contains(serviceNameNormal), "Le service " + serviceNameNormal + " n'existe pas.");
            Assert.IsTrue(serviceNormalMAD.Contains(serviceNameNormalMAD), "Le service " + serviceNameNormalMAD + " n'existe pas.");
            Assert.IsTrue(serviceVacuum.Contains(serviceNameVacuum), "Le service " + serviceNameVacuum + " n'existe pas.");
            Assert.IsTrue(serviceFlightValide.Contains(serviceNameFlightValide), "Le service " + serviceNameFlightValide + " n'existe pas.");
            Assert.IsTrue(serviceSansDatasheet.Contains(serviceNameSansDatasheet), "Le service " + serviceNameSansDatasheet + " n'existe pas.");
        }

        // _____________________________________________ Filtres et page Filter And Favorite__________________________________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_Site()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameNormalMAD = TestContext.Properties["Prodman_ServiceNormalMAD"].ToString();

            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameNormalMAD = TestContext.Properties["Prodman_FlightNormalMAD"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteMAD);
            bool IsServiceMAD = qtyAjustementPage.IsServicePresent(serviceNameNormalMAD);

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            bool IsServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            // Vérification des services pour les différents sites
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + "n'est pas présent pour le site " + siteACE);
            Assert.IsFalse(qtyAjustementPage.IsServicePresent(serviceNameNormalMAD), "Le service " + serviceNameNormalMAD + "est présent pour le site " + siteACE);

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteMAD);
            Assert.IsFalse(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + "est présent pour le site " + siteMAD);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormalMAD), "Le service " + serviceNameNormal + "n'est pas présent pour le site " + siteMAD);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_Site_FaF()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // On filtre les données par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, site);
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            string siteSelected = qtyAjustementPage.GetSite();
            Assert.AreEqual(site, siteSelected, String.Format(MessageErreur.FILTRE_ERRONE, "Le site n est pas selectionner sur l interface qty "));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_SortBy()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();

            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string customerName2 = TestContext.Properties["Prodman_Customer2"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameVacuum = TestContext.Properties["Prodman_ServiceVacuum"].ToString();

            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameVacuum = TestContext.Properties["Prodman_FlightVacuum"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool IsServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool IsServiceVacuum = qtyAjustementPage.IsServicePresent(serviceNameVacuum);

            if (!IsServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);
            }

            if (!IsServiceVacuum)
            {
                CreateFlight(homePage, siteACE, flightNameVacuum, customerName2, guestType2, serviceNameVacuum);
            }

            if (!IsServiceNormal || !IsServiceVacuum)
            {
                filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            }

            // Sort by customer
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.SortBy, "Customer");
            bool isSortedByCustomer = qtyAjustementPage.IsSortedByCustomer();

            // Sort by service
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.SortBy, "Service");
            qtyAjustementPage.PageSize("100");
            bool isSortedByService = qtyAjustementPage.IsSortedByService();

            //assert
            Assert.IsTrue(isSortedByCustomer, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by customers"));
            Assert.IsTrue(isSortedByService, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by service"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_FromTo()
        {
            // Pour ce test, nous avons besoin de deux services : normal (un service J+1 et un serivce J-2 ) .
            // L'objectif de ce test est de tester la fonctionnalité du filtre From/To sur le service Normal. C'est pour cela que nous avons créé 2 services
            // Avec des dates != 
            // Cela nous permettra d'obtenir le résultat du filtre sur Normal .

            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();

            DateTime dateFrom = DateUtils.Now.AddDays(-2);
            DateTime dateTo = DateUtils.Now;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.PageSize("100");
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            // Assert 1 : verif qu'un service est disponible avec Jour J 
            Assert.IsTrue(isServiceNormal, "pas de service Normal ");

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateFrom, dateFrom);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, dateFrom);
            bool isDateFromOkQtyAdjBefore2Days = qtyAjustementPage.VerifyDateFrom(dateFrom);
            bool isDateToOkQtyAdjBefore2Days = qtyAjustementPage.VerifyDateTo(dateFrom);
            // Assert 1 : verif mis en palce Date avant 2 J 
            Assert.IsTrue(isDateFromOkQtyAdjBefore2Days, String.Format(MessageErreur.FILTRE_ERRONE, "From Qty adjustement"));
            Assert.IsTrue(isDateToOkQtyAdjBefore2Days, String.Format(MessageErreur.FILTRE_ERRONE, "To Qty adjustement"));
            // Assert 2 : verif qu'un service est disponible avec J-2 
            bool isServiceNormalAvant = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(isServiceNormalAvant, "pas de service avant 2 jours ");

            // Page qty adjustement
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, dateTo);
            qtyAjustementPage.ClickTabQtyAdj();
            bool isDateFromOkQtyAdj = qtyAjustementPage.VerifyDateFrom(dateFrom);
            bool isDateToOkQtyAdj = qtyAjustementPage.VerifyDateTo(dateTo);
            isServiceNormalAvant = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            // Assert : Date 
            Assert.IsTrue(isDateFromOkQtyAdj, String.Format(MessageErreur.FILTRE_ERRONE, "From Qty adjustement"));
            Assert.IsTrue(isDateToOkQtyAdj, String.Format(MessageErreur.FILTRE_ERRONE, "To Qty adjustement"));
            // Assert : verif service est valide sur un plage de date 
            Assert.IsTrue(isServiceNormalAvant, "pas de service avant 2 jours ");
            // Page results
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            bool isWorkShopDisplayed = resultPage.IsCorrectRawMatByWorkshopSelected(); 
            bool isDateFromOkResult = resultPage.VerifyDateFrom(dateFrom.Date);
            bool isDateToOkResult = resultPage.VerifyDateTo(dateTo.Date);
            // Assert
            Assert.IsTrue(isWorkShopDisplayed, "WorkShop n est pas visible . "); 
            Assert.IsTrue(isDateFromOkResult, String.Format(MessageErreur.FILTRE_ERRONE, "From Results"));
            Assert.IsTrue(isDateToOkResult, String.Format(MessageErreur.FILTRE_ERRONE, "To Results"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_FromTo_FaF()
        {
            // Prepare
            DateTime dateFrom = DateUtils.Now.AddDays(-2);
            DateTime dateTo = DateUtils.Now.AddDays(+1);

            //Arrange
            HomePage homePage = LogInAsAdmin();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // On filtre par production Date
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateFrom, dateFrom);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateTo, dateTo);
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            // On récupère les dates dans la page qty adjustement
            DateTime dateFromFilter = qtyAjustementPage.GetDateFrom();
            DateTime dateToFilter = qtyAjustementPage.GetDateTo();

            Assert.AreEqual(dateFrom.Date, dateFromFilter.Date, String.Format(MessageErreur.FILTRE_ERRONE, "From FaF"));
            Assert.AreEqual(dateTo.Date, dateToFilter.Date, String.Format(MessageErreur.FILTRE_ERRONE, "To FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_BatchStartTime()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Raw materials Report_-_";
            string DocFileNameZipBegin = "All_files_";
            // Arrange 
            HomePage homePage = LogInAsAdmin();
            homePage.ClearDownloads();
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.PageSize("100");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "pas de service normal disponible");
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.StartTime, "22:30");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), String.Format(MessageErreur.FILTRE_ERRONE, "Start Time & End Time"));


            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            PrintReportPage reportPage = resultPage.PrintReport(ResultPage.PrintType.RawMaterialsByWorkshop, true);
            reportPage.Close();
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            string reconstructedText = string.Join(" ", words.Select(w => w.Text));
            string expectedText = " Start Time 22:30 "; 
            int startIndex = reconstructedText.IndexOf(expectedText);
            Assert.AreNotEqual(-1, startIndex, "Start Time non visible "); 
            string extractedText = reconstructedText.Substring(startIndex, 16);
            Assert.AreEqual(extractedText, expectedText, "Start time is wrong or not displayed in pdf document Raw materials");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_BatchEndTime()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Raw materials Report_-_";
            string DocFileNameZipBegin = "All_files_";
            // Arrange 
            HomePage homePage = LogInAsAdmin();
            homePage.ClearDownloads();
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.PageSize("100");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "pas de service normal disponible");
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.EndTime, "23:30");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), String.Format(MessageErreur.FILTRE_ERRONE, "Start Time & End Time"));


            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            PrintReportPage reportPage = resultPage.PrintReport(ResultPage.PrintType.RawMaterialsByWorkshop, true);
            reportPage.Close();
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            string reconstructedText = string.Join(" ", words.Select(w => w.Text));
            string expectedText = " End Time 23:30 ";
            int startIndex = reconstructedText.IndexOf(expectedText);
            string extractedText = reconstructedText.Substring(startIndex  ,16 );
            Assert.AreEqual(extractedText , expectedText , "End time is wrong or not displayed in pdf document Raw materials");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ShowWithoutDatasheet()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameSansDatasheet = TestContext.Properties["Prodman_ServiceSansDatasheet"].ToString();
            string flightNameSansDatasheet = TestContext.Properties["Prodman_FlightSansDatasheet"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, DateUtils.Now);
            bool isServiceSansDatasheet = qtyAjustementPage.IsServicePresent(serviceNameSansDatasheet);

            if (!isServiceSansDatasheet)
            {
                CreateFlight(homePage, siteACE, flightNameSansDatasheet, customerName1, guestType2, serviceNameSansDatasheet);

                filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, DateUtils.Now);
            }

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowWithoutDatasheet, true);
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

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ShowValidatedFlightsOnly()
        {
            // Pour ce test, nous avons besoin de deux vols : "pré-validé" et "validé".
            // L'objectif de ce test est de vérifier la fonctionnalité ShowValidatedFlightsOnly  sur l'interface des résultats.
            // Ainsi, après l'application du filtre, nous devrions voir uniquement le service lié aux vols validés, et non aux vols pré-validés.

            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();


            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, DateUtils.Now);
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            Assert.IsTrue(isServiceNormal, "Le service " + serviceNameNormal + " n'est pas présent.");
            Assert.IsTrue(isServiceValide, "Le service " + serviceNameFlightValide + " n'est pas présent.");

            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowValidateFlight, true);
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);

            resultPage.GoToQtyAdjustementPage();

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, DateUtils.Now);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameFlightValide), "Le service " + serviceNameFlightValide + " n'est pas présent.");
            Assert.IsFalse(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas présent.");

            // Page results
            resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            HashSet<string> flights = resultPage.GetFlights();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.SetDateState(DateUtils.Now.AddDays(1));
            string valid = "Valid";
            flightPage.Filter(FlightPage.FilterType.Status, valid);
            // Assert sur qu on a des vols 
            Assert.IsTrue(flights.Count > 0, "No flights to search for ."); 
            foreach (string flightName in flights)
            {
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightName);
                Assert.IsTrue(flightPage.GetFirstFlightStatus("V"), String.Format(MessageErreur.FILTRE_ERRONE, "Show validated flight only"));
            }
        }

        [TestMethod]
        [Ignore] // supprime le premier article mais ne le rétabli pas, effet de bord sur tous les tests production
        public void PR_PRODMAN_Filter_ShowHiddenArticle()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();

            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = TestContext.Properties["Prodman_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

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
                filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
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

            // On cache le premier article
            resultPage.HideFirstArticle();

            filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            items = resultPage.GetItemNames();
            int nbItemFinal = resultPage.CountItems();

            //assert
            Assert.IsFalse(items.Contains(firstItemName), "L'item " + firstItemName + " n'a pas été caché.");
            Assert.AreNotEqual(nbItemInit, nbItemFinal, "le nombre d'items visible n'a pas été modifié.");


            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowHiddenArticles, true);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            items = resultPage.GetItemNames();
            Assert.IsTrue(items.Contains(firstItemName), String.Format(MessageErreur.FILTRE_ERRONE, "Show hidden article"));
            Assert.AreEqual(nbItemInit, resultPage.CountItems(), "Le nombre d'items visibles n'a pas été retabli malgré l'utilisation du filtre.");

            // On fait réapparaitre le premier item
            resultPage.HideFirstArticle();

            filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            items = resultPage.GetItemNames();

            //assert
            Assert.IsTrue(items.Contains(firstItemName), "L'item " + firstItemName + " n'a pas été retabli.");
            Assert.AreEqual(nbItemInit, resultPage.CountItems(), "le nombre d'items visible n'a pas été rétabli.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ShowNormalAndVacuumProduction()
        {
            // Pour ce test, nous avons besoin de deux service lié a 2  vol (NORMAL ET VACCUM) .
            // L'objectif de ce test est de verifier que l option d afficher seulement service normal et vaccum est fonctionnel  .
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameVacuum = TestContext.Properties["Prodman_ServiceVacuum"].ToString();
      
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, DateUtils.Now);
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceVacuum = qtyAjustementPage.IsServicePresent(serviceNameVacuum);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameVacuum), "Le service " + serviceNameVacuum + " n'est pas visible.");
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowVacuumProd, true);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameVacuum), "Le service " + serviceNameVacuum + " n'est pas visible.");
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowNormalProd, true);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowNormalAndVacuumProd, true);
            List<string> servicesName = qtyAjustementPage.GetServicesName();
            Assert.IsTrue(servicesName.Contains(serviceNameNormal) && servicesName.Contains(serviceNameVacuum) , "le filtre n est pas pris en compte ");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ShowNormalAndVacuumProduction_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "show normal and vacuum" et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.ShowNormalAndVacuumProd, true);

            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            //assert
            Assert.IsTrue(qtyAjustementPage.IsNormalAndVacuumProduction(), String.Format(MessageErreur.FILTRE_ERRONE, "Show normal and vacuum production FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ShowVacuumProductionOnly()
        {
            // Pour ce test nous avons besoin de deux service normal et vaccum 
            // L'objectif de ce test est de verifier la fonct° d afficher seulement le service Vaccum 
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameVacuum = TestContext.Properties["Prodman_ServiceVacuum"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");

            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceVacuum = qtyAjustementPage.IsServicePresent(serviceNameVacuum);

            Assert.IsTrue(isServiceNormal, "Le service " + serviceNameNormal + " n'est pas visible.");
            Assert.IsTrue(isServiceVacuum, "Le service " + serviceNameVacuum + " n'est pas visible.");

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowVacuumProd, true);
            isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            isServiceVacuum = qtyAjustementPage.IsServicePresent(serviceNameVacuum);

            Assert.IsFalse(isServiceNormal, "Le service " + serviceNameNormal + " n'est pas visible.");
            Assert.IsTrue(isServiceVacuum, "Le service " + serviceNameVacuum + " n'est pas visible.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ShowVacuumProductionOnly_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "show vacuum only" et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.ShowVacuumProd, true);

            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            //assert
            Assert.IsTrue(qtyAjustementPage.IsVacuumProduction(), String.Format(MessageErreur.FILTRE_ERRONE, "Show vacuum production only FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ShowNormalProductionOnly()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string recipeCookingMode = TestContext.Properties["Prodman_Needs_CookingMode2"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string customerName2 = TestContext.Properties["Prodman_Customer2"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameVacuum = TestContext.Properties["Prodman_ServiceVacuum"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameVacuum = TestContext.Properties["Prodman_FlightVacuum"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            bool isServiceNameNormalPresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(isServiceNameNormalPresent, $"Le service {serviceNameNormal} n'est pas visible.");
            bool isServiceNameVacuumPresent = qtyAjustementPage.IsServicePresent(serviceNameVacuum);
            Assert.IsTrue(isServiceNameVacuumPresent, $"Le service {serviceNameVacuum} n'est pas visible.");
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowNormalProd, true);
            Assert.IsTrue(isServiceNameNormalPresent, $"Le service {serviceNameNormal} n'est pas visible.");
            isServiceNameVacuumPresent = qtyAjustementPage.IsServicePresent(serviceNameVacuum);
            Assert.IsFalse(isServiceNameVacuumPresent, $"Le service {serviceNameVacuum} est visible.");
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.WaitLoading();
            resultPage.UnfoldAll();
            List<string> cookingModesGroup = resultPage.GetCookingModes(ResultPage.CookingModeType.COOKING_MODE_GROUP);
            bool isNormalOKGroup = (cookingModesGroup.Count != 0 && !cookingModesGroup.Contains(recipeCookingMode));
            Assert.IsTrue(isNormalOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal production only Raw mat by group"));
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.WaitLoading();
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();
            List<string> cookingModesWorkshop = resultPage.GetCookingModes(ResultPage.CookingModeType.COOKING_MODE_WORKSHOP);
            bool isNormalOKWorkshop = (cookingModesWorkshop.Count != 0 && !cookingModesWorkshop.Contains(recipeCookingMode));
            Assert.IsTrue(isNormalOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal production only Raw mat by Workshop"));
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.WaitLoading();
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();
            List<string> cookingModesSupplier = resultPage.GetCookingModes(ResultPage.CookingModeType.COOKING_MODE_SUPPLIER);
            bool isNormalOKSupplier = (cookingModesSupplier.Count != 0 && !cookingModesSupplier.Contains(recipeCookingMode));
            Assert.IsTrue(isNormalOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal production only Raw mat by Supplier"));
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.WaitLoading();
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();
            List<string> cookingModesRecipe = resultPage.GetCookingModes(ResultPage.CookingModeType.COOKING_MODE_RECIPE);
            bool isNormalOKRecipe = cookingModesRecipe.Count != 0 && !cookingModesRecipe.Contains(recipeCookingMode);
            Assert.IsTrue(isNormalOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Show normal production only Raw mat by Recipe"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ShowNormalProductionOnly_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "show normal only" et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.ShowNormalProd, true);

            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            //assert
            Assert.IsTrue(qtyAjustementPage.IsNormalProduction(), String.Format(MessageErreur.FILTRE_ERRONE, "Show normal production only FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_Workshops()
        {
            // Pour ce test, nous avons besoin de deux services : normal et vacuum.
            // L'objectif de ce test est de vérifier la fonctionnalité du filtre workshop.
            // Nous prenons le workshop du service normal et du service vacuum, puis nous filtrons en nous basant sur ce workshop retenu.
            // Nous validons les workshops sur le datasheet (chaque item doit avoir ce workshop).
            // !!! Attention aux articles hidden !!!

            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();
   
            //Arrange
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, DateUtils.Now);
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            Assert.IsTrue(isServiceNormal, "pas de service Normal ");
            Assert.IsTrue(isServiceValide, "pas de service Valide");

            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.GroupBy(GROUP_BY_WORKSHOP);
            string workshop = resultPage.GetFirstItemGroup();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Workshops, workshop);

            resultPage.FilterRawBy(RAW_MAT_GROUP);
            string workshopReturned = resultPage.GetFirstItemGroup();
            resultPage.UnfoldAll();
            bool isWorkshopOKGroup = resultPage.VerifyWorkshop(workshop, true);
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();
            bool isWorkshopOKWorkshop = resultPage.VerifyWorkshop(workshop);
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();
            bool isWorkshopOKSupplier = resultPage.VerifyWorkshop(workshop);
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();
            bool isWorkshopOKRecipe = resultPage.VerifyWorkshop(workshop);

            Assert.AreEqual(workshop, workshopReturned, "Le nom du workshop ne correspond pas à celui filtré dans 'Raw mat by group'.");
            Assert.IsTrue(isWorkshopOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Workshop Raw mat by group"));
            Assert.IsTrue(isWorkshopOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Workshop Raw mat by Workshop"));
            Assert.IsTrue(isWorkshopOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Workshop Raw mat by Supplier"));
            Assert.IsTrue(isWorkshopOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Workshop Raw mat by Recipe"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_Workshops_FaF()
        {
            // Prepare
            string recipeWorkshop1 = TestContext.Properties["Prodman_Needs_Workshop1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            // Filtre par workshop
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Workshops, recipeWorkshop1);
            // Verification sur Qty 
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            string selectedWorkshop = qtyAjustementPage.GetWorkshop();
            //assert
            Assert.AreEqual(recipeWorkshop1, selectedWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Workshop n est pas pris en charge "));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_Customers()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();


            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");

            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            Assert.IsTrue(isServiceNormal, "pas de service Normal ");
            Assert.IsTrue(isServiceValide, "pas de service Valide");

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, DateUtils.Now.AddDays(1));
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Customers, customerName3);
            //qtyAjustementPage.PageSize("100");

            // Page Qty ajustement
            bool isCustomerOKQtyAjus = qtyAjustementPage.VerifyCustomer(customerName3);

            // Page results
            // ___________________ RAW MAT BY GROUP ______________________________
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
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

            bool isCustomerOKWorkshop = resultPage.VerifyCustomerRawMat(customerName3);

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
            bool isCustomerOKRecipe = resultPage.VerifyCustomerRawMat(customerName3);

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
        public void PR_PRODMAN_Filter_Customers_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtrer par client et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Customers, customerName1);
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            // Vérifier si le client "customerName1" a bien été sélectionné dans le filtre
            string selectedCustomer = qtyAjustementPage.GetCustomer();
            // Assert1 : La liste de client contain le customerName1
            Assert.IsTrue(selectedCustomer.Contains(customerName1), String.Format(MessageErreur.FILTRE_ERRONE, "Le client n'est pas sélectionné"));
            // Assert 2 : 1 of X customers selected 
            string filterCustomer = qtyAjustementPage.GetCustomerFilterNumber();
            string firstCharacterFromTheCustomer = filterCustomer[0].ToString();
            Assert.AreEqual(firstCharacterFromTheCustomer, "1", "the filter selected multipale customers ");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_GuestType()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();

            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = TestContext.Properties["Prodman_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, DateUtils.Now);
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            Assert.IsTrue(isServiceNormal, "pas de service Normal ");
            Assert.IsTrue(isServiceValide, "pas de service de nom Valide");

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.GuestType, guestType1);

            // Page qty ajustements
            bool isServiceNormalPresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValidePresent = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            Assert.IsTrue(isServiceNormalPresent == true && isServiceValidePresent == true, String.Format(MessageErreur.FILTRE_ERRONE, "Guest type qty adjustements"));

            // Page results
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.FilterRawBy(RAW_MAT_CUSTOMER);
            resultPage.UnfoldAll();
            resultPage.WaitPageLoading();
            bool returnGuestType = resultPage.VerifyGuestType(guestType1);

            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            HashSet<string> flights = resultPage.GetFlights();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteACE);
            flightPage.SetDateState(DateUtils.Now.AddDays(+1));

            foreach (string flightName in flights)
            {
                if (flightName != "")
                {
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightName);
                    // Find the index of the first date (after the "- ")
                    int startIndex = flightName.IndexOf("-") + 2;  // "+ 2" to skip "- "

                    // Extract the first date
                    string firstDate = flightName.Substring(startIndex, 10);
                    flightPage.SetDateState(DateTime.ParseExact(firstDate, "dd/MM/yyyy", CultureInfo.InvariantCulture));
                    FlightDetailsPage editPage = flightPage.EditFirstFlight(flightName);
                    string guestRetourne = editPage.GetGuestType();
                    flightPage = editPage.CloseViewDetails();
                    Assert.AreEqual(guestType1, guestRetourne, String.Format(MessageErreur.FILTRE_ERRONE, "Guest type"));

                }

            }

            Assert.IsTrue(returnGuestType, String.Format(MessageErreur.FILTRE_ERRONE, "Guest type results"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_GuestType_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            // Appliquer les filltres via l'interface filtre et favorites 
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par guestType et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.GuestType, guestType1);

            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            // Verfier le choie de filtre selectionner 
            string selectedGuestType = qtyAjustementPage.GetGuestType();
            //assert
            Assert.AreEqual(guestType1, selectedGuestType, String.Format(MessageErreur.FILTRE_ERRONE, "Guest type FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ServiceCategorie()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();
            string serviceCategorie1 = TestContext.Properties["Prodman_Needs_ServiceCategory1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");

            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameFlightValide), "Le service " + serviceNameFlightValide + " n'est pas visible.");

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ServicesCategorie, serviceCategorie1);

            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            // ___________________ RAW MAT BY GROUP ______________________________
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            bool isCategorieOKGroup = resultPage.VerifyCategorie(serviceCategorie1);

            // ___________________ RAW MAT WORKSHOP ______________________________
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isCategorieOKWorkshop = resultPage.VerifyCategorie(serviceCategorie1);

            // ___________________ RAW MAT SUPPLIER ______________________________
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isCategorieOKSupplier = resultPage.VerifyCategorie(serviceCategorie1);

            // ___________________ RAW MAT RECIPE ______________________________
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isCategorieOKRecipe = resultPage.VerifyCategorie(serviceCategorie1);

            //assert
            Assert.IsTrue(isCategorieOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Service categorie raw mat by group"));
            Assert.IsTrue(isCategorieOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Service categorie raw mat by workshop"));
            Assert.IsTrue(isCategorieOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Service categorie raw mat by supplier"));
            Assert.IsTrue(isCategorieOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Service categorie raw mat by recipe"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ServiceCategorie_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceCategorie1 = TestContext.Properties["Prodman_Needs_ServiceCategory1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par guestType et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.ServicesCategorie, serviceCategorie1);

            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            string selectedServiceCategorie = qtyAjustementPage.GetServiceCategorie();

            //assert
            Assert.AreEqual(serviceCategorie1, selectedServiceCategorie, String.Format(MessageErreur.FILTRE_ERRONE, "Service categorie FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_Service()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");

            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            Assert.IsTrue(isServiceNormal, "pas de service Normal ");
            Assert.IsTrue(isServiceValide, "pas de service de nom Valide");

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameNormal);
            //qtyAjustementPage.PageSize("100");


            // page qty adjustement
            Assert.AreEqual(1, qtyAjustementPage.CountServices(), "Il y a plusieurs services visibles malgré le filtre.");
            string returnService = qtyAjustementPage.GetServiceName();

            // Page results
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            // ___________________ RAW MAT BY GROUP ______________________________
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            bool isServiceOKGroup = resultPage.VerifyService(serviceNameNormal);

            // ___________________ RAW MAT WORKSHOP ______________________________
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isServiceOKWorkshop = resultPage.VerifyService(serviceNameNormal);

            // ___________________ RAW MAT SUPPLIER ______________________________
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isServiceOKSupplier = resultPage.VerifyService(serviceNameNormal);

            // ___________________ RAW MAT RECIPE ______________________________
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isServiceOKRecipe = resultPage.VerifyService(serviceNameNormal);

            //assert
            Assert.AreEqual(returnService, serviceNameNormal, string.Format(MessageErreur.FILTRE_ERRONE, "Filter Service qty adjustement"));
            Assert.IsTrue(isServiceOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Service raw mat by group"));
            Assert.IsTrue(isServiceOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Service raw mat by workshop"));
            Assert.IsTrue(isServiceOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Service raw mat by supplier"));
            Assert.IsTrue(isServiceOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Service raw mat by recipe"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_Service_FaF()
        {
            // Prepare
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par guestType et par site          
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Service, serviceNameNormal);

            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            string selectedService = qtyAjustementPage.GetService();

            //assert
            Assert.AreEqual(serviceNameNormal, selectedService, String.Format(MessageErreur.FILTRE_ERRONE, "Service FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_RecipeType()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();

            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = TestContext.Properties["Prodman_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            int offSetDay = 1;
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");

            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, DateUtils.Now);
            while (qtyAjustementPage.CountServices() == 0)
            {
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateFrom, DateUtils.Now.AddDays(-offSetDay));
                offSetDay++;
            }
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            // assert de service 
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.GroupbyRawMatGroup(GROUP_BY_RECIPE_TYPE);
            string recipeType = resultPage.GetFirstItemGroup();

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.RecipeType, recipeType);

            // ___________________ RAW MAT BY GROUP ______________________________
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            string recipeTypeReturned = resultPage.GetFirstItemGroup();
            resultPage.UnfoldAll();

            bool isRecipeTypeOKGroup = resultPage.VerifyRecipeType(recipeType);

            // ___________________ RAW MAT WORKSHOP ______________________________
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isRecipeTypeOKWorkshop = resultPage.VerifyRecipeType(recipeType);

            // ___________________ RAW MAT SUPPLIER ______________________________
            resultPage.FilterRawBy(RAW_MAT_SUPPLIER);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isRecipeTypeOKSupplier = resultPage.VerifyRecipeType(recipeType);

            // ___________________ RAW MAT RECIPE ______________________________
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();
            resultPage.UnfoldAll();

            bool isRecipeTypeOKRecipe = resultPage.VerifyRecipeType(recipeType);


            Assert.AreEqual(recipeType, recipeTypeReturned, "Le nom du recipe type ne correspond pas à celui filtré dans 'Raw mat by group'.");
            Assert.IsTrue(isRecipeTypeOKGroup, String.Format(MessageErreur.FILTRE_ERRONE, "Recipe type Raw mat by group"));
            Assert.IsTrue(isRecipeTypeOKWorkshop, String.Format(MessageErreur.FILTRE_ERRONE, "Recipe type Raw mat by Workshop"));
            Assert.IsTrue(isRecipeTypeOKSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "Recipe type Raw mat by Supplier"));
            Assert.IsTrue(isRecipeTypeOKRecipe, String.Format(MessageErreur.FILTRE_ERRONE, "Recipe type Raw mat by Recipe"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_RecipeType_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string recipeType1 = TestContext.Properties["Prodman_Needs_RecipeType1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par guestType et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RecipeType, recipeType1);

            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            string selectedRecipeType = qtyAjustementPage.GetRecipeType();

            //assert
            Assert.AreEqual(recipeType1, selectedRecipeType, String.Format(MessageErreur.FILTRE_ERRONE, "RecipeType FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ItemGroup()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            Assert.AreEqual("All", filterAndFavorite.CheckTAllItemGroupsSelected(), "Tous les item Groups ne sont pas sélectionnés en arrivant sur la page : " + filterAndFavorite.CheckTAllItemGroupsSelected());
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.GroupbyRawMatGroup(GROUP_BY_ITEMGROUP);
            Assert.AreNotEqual(0, resultPage.CountResults(), "Il n'y a pas de groupes d'items visibles.");
            var firstItemGroup = resultPage.GetFirstItemGroup();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ItemGroups, firstItemGroup);
            Assert.AreEqual(1, resultPage.CountResults(), "Il plusieurs groupes visibles après application du filtre ItemGroups.");
            Assert.AreEqual(firstItemGroup, resultPage.GetFirstItemGroup(), string.Format(MessageErreur.FILTRE_ERRONE, "ItemGroups"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ItemGroup_FaF()
        {

            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string itemGroup1 = TestContext.Properties["Prodman_Needs_Item_Group1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par guestType et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);

            // Assert 1 : Quand on sélectionne un site, tous les item groups doivent être sélectionnés
            Assert.AreEqual("All", filterAndFavorite.CheckTAllItemGroupsSelected(), "Tous les item Groups ne sont pas sélectionnés après la sélection d'un site : " + filterAndFavorite.CheckTAllItemGroupsSelected());

            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.ItemGroups, itemGroup1);

            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            string selectedItemGroup = qtyAjustementPage.GetItemGroup();
            //Assert 2 : Vérifier l'application correct de filtre 
            Assert.AreEqual(itemGroup1, selectedItemGroup, String.Format(MessageErreur.FILTRE_ERRONE, "ItemGroup FaF"));

        }

        // __________________________________________ Page Filter and favorite ___________________________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_MakeFavorite()
        {
            // On fait un vil avec ServiceNameNormal , customerName1 , siteACE ---> on utilse cette vol pour un filtre 

            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string group1 = TestContext.Properties["Prodman_Needs_Item_Group1"].ToString();
            string favoriteName = "testProdman";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Paramétrage et filtrage d'un favori
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Service, serviceNameNormal);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Customers, customerName1);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.ItemGroups, group1);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.ShowNormalProd, true);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RawMaterialByWorkshop, true);

            // Creartion de favorit 
            filterAndFavorite.MakeFavorite(favoriteName);
            filterAndFavorite.SetFavoriteText(favoriteName);
            // Assert 1 : Le favoris est ajouté
            Assert.IsTrue(filterAndFavorite.IsFavoritePresent(favoriteName), "Le favori " + favoriteName + " n'a pas été créé.");

            try
            {
                filterAndFavorite.ResetFilter();
                filterAndFavorite.SetFavoriteText(favoriteName);
                // On active le favorite
                filterAndFavorite.SelectFavorite(favoriteName);
                // Assert 2 : La verification de conservation de filtre 
                Assert.IsTrue(filterAndFavorite.GetNbService().Contains("1 of "), "Le filtre Service du favori n'a pas été conservé.");
                Assert.IsTrue(filterAndFavorite.GetNbItemGroup().Contains("1 of "), "Le filtre ItemGroup du favori n'a pas été conservé.");
                Assert.IsTrue(filterAndFavorite.IsNormalProduction(), "Le filtre 'Show normal production only' du favori n'a pas été conservé.");
                Assert.IsTrue(filterAndFavorite.IsRawMatByWorkshop(), "Le filtre 'Raw mat by workshop' du favori n'a pas été conservé.");
            }
            finally
            {
                // suppression de favorite cree 
                filterAndFavorite.SetFavoriteText(favoriteName);
                filterAndFavorite.DeleteFavorite(favoriteName);
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ItemGroup_QA()
        {
            // Pour ce test, nous avons besoin d'un service name Normal.
            // L'objectif de ce test est de vérifier la fonctionnalité du filtre ItemGroup.
            // Nous avons besoin de prendre le nom de l'ItemGroup depuis l'interface Results.

            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.GroupbyRawMatGroup(GROUP_BY_ITEMGROUP);

            Assert.AreNotEqual(0, resultPage.CountResults(), "Il n'y a pas de groupes d'items visibles.");

            string firstItemGroup = resultPage.GetFirstItemGroup();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ItemGroups, firstItemGroup);

            Assert.AreEqual(1, resultPage.CountResults(), "Il plusieurs groupes visibles après application du filtre ItemGroups.");
            Assert.AreEqual(firstItemGroup, resultPage.GetFirstItemGroup(), string.Format(MessageErreur.FILTRE_ERRONE, "ItemGroups"));
        }
        // __________________________________________ Page Quantity adjustement __________________________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Edit_Qty_Adjustement()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            // On modifie la quantité du premier service visible
            var initQty = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal);
            qtyAjustementPage.SelectService(serviceNameNormal);
            qtyAjustementPage.SetQuantity((initQty + 1).ToString(), serviceNameNormal);
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            qtyAjustementPage = resultPage.GoToQtyAdjustementPage();

            //Assert
            qtyAjustementPage.SelectService(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsQuantityUpdated((initQty + 1).ToString(), serviceNameNormal), "La modification n'est pas affichée.");

            qtyAjustementPage.ResetAllQuantities();
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_ResetAllQuantities()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            string decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            Assert.IsTrue(isServiceNormal, "le service nnormal nest pas existant");
            Assert.IsTrue(isServiceValide, "le service flight valide n'est pas existant");

            // On modifie les quantités de 2 services visibles
            double initQtyNormal = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal);
            qtyAjustementPage.SetQuantityWithSelect((initQtyNormal + 22).ToString(), serviceNameNormal);

            double initQtyValide = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameFlightValide);
            qtyAjustementPage.SetQuantityWithSelect((initQtyValide + 22).ToString(), serviceNameFlightValide);

            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            qtyAjustementPage = resultPage.GoToQtyAdjustementPage();

            // On vérifie que les quantités ont été modifiées
            Assert.AreEqual((initQtyNormal + 22), qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal), "La modification n'est pas effectuée sur le service " + serviceNameNormal);
            Assert.AreEqual((initQtyValide + 22), qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameFlightValide), "La modification n'est pas effectuée sur le service " + serviceNameFlightValide);

            // On reset les quantités
            qtyAjustementPage.ResetAllQuantities();

            //Assert
            Assert.AreEqual(initQtyNormal, qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal), "Le reset n'a pas fonctionné sur la quantité du service " + serviceNameNormal);
            Assert.AreEqual(initQtyValide, qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameFlightValide), "Le reset n'a pas fonctionné sur la quantité du service " + serviceNameFlightValide);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_RazAllQuantities()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            string decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            Assert.IsTrue(isServiceNormal, "le service nnormal nest pas existant");
            Assert.IsTrue(isServiceValide, "le service flight valide n'est pas existant");

            // On modifie les quantités de 2 services visibles
            double initQtyNormal = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal);

            if (initQtyNormal == 0)
            {
                qtyAjustementPage.SetQuantityWithSelect("1", serviceNameNormal);
            }

            double initQtyValide = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameFlightValide);

            if (initQtyValide == 0)
            {
                qtyAjustementPage.SetQuantityWithSelect("1", serviceNameFlightValide);
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
        public void PR_PRODMAN_NextOrPreviousDate()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

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
        public void PR_PRODMAN_RawMatByGroup()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            // ajout assert 
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
        public void PR_PRODMAN_RawMatByGroup_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "Raw mat by group" et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);

            // Group by ItemGroup
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RawMaterialByGroup, true);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.GroupBy, GROUP_BY_ITEMGROUP);

            ResultPage resultsPage = filterAndFavorite.DoneToResults();
            Assert.IsTrue(resultsPage.GetRawMatByPage(RAW_MAT_GROUP) && resultsPage.GetGroupByPage(GROUP_BY_ITEMGROUP), String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by group-- ItemGroup FaF"));

            // Group by Workshop
            filterAndFavorite = resultsPage.Back();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RawMaterialByGroup, true);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.GroupBy, GROUP_BY_WORKSHOP);

            resultsPage = filterAndFavorite.DoneToResults();
            Assert.IsTrue(resultsPage.GetRawMatByPage(RAW_MAT_GROUP) && resultsPage.GetGroupByPage(GROUP_BY_WORKSHOP), String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by group-- Workshop FaF"));

            // Group by customer
            filterAndFavorite = resultsPage.Back();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RawMaterialByGroup, true);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.GroupBy, GROUP_BY_CUSTOMER);

            resultsPage = filterAndFavorite.DoneToResults();
            Assert.IsTrue(resultsPage.GetRawMatByPage(RAW_MAT_GROUP) && resultsPage.GetGroupByPage(GROUP_BY_CUSTOMER), String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by group-- Customer FaF"));

            // Group by recipe type
            filterAndFavorite = resultsPage.Back();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RawMaterialByGroup, true);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.GroupBy, GROUP_BY_RECIPE_TYPE);

            resultsPage = filterAndFavorite.DoneToResults();
            Assert.IsTrue(resultsPage.GetRawMatByPage(RAW_MAT_GROUP) && resultsPage.GetGroupByPage(GROUP_BY_RECIPE_TYPE), String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by group-- Recipe type FaF"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_RawMatByWorkhop()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            Assert.IsTrue(isServiceNormal, "le service nnormal nest pas existant");
            Assert.IsTrue(isServiceValide, "le service flight valide n'est pas existant");
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);
            resultPage.WaitPageLoading();
            resultPage.UnfoldAll();
            resultPage.WaitPageLoading();

            bool hasSubgroupAndRecipeAssociated = resultPage.GetSubgroupAndRecipeAssociated();
            Assert.IsTrue(hasSubgroupAndRecipeAssociated, "Les informations affichées dans la vue 'Raw mat by workshop' ne sont pas correctes.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_RawMatByWorkshop_FaF()
        {


            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            // On a cree un vol avec customerName1 qui est associer avec WorkShop1
            string WorkShop1 = TestContext.Properties["Prodman_Needs_Workshop1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "Raw mat by workshop" et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RawMaterialByWorkshop, true);
            ResultPage resultsPage = filterAndFavorite.DoneToResults();
            bool isCorrectRawMatByWorkshop = resultsPage.IsCorrectRawMatByWorkshopSelected();

            // Assert
            Assert.IsTrue(isCorrectRawMatByWorkshop, "les résultats ne s'accordent pas au filtre");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_RawMatBySupplier()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();

            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = TestContext.Properties["Prodman_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            if (!isServiceNormal || !isServiceValide)
            {
                filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
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
        public void PR_PRODMAN_RawMatBySupplier_FaF()
        {
            // Prepare 
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtrer par "Raw mat by supplier" et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RawMaterialBySupplier, true);
            ResultPage resultsPage = filterAndFavorite.DoneToResults();

            // Filtrer pour afficher les données cachées
            resultsPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);

            // Assert 1 : vérifier que la vue des résultats a bien sélectionné "Raw mat by supplier"
            Assert.IsTrue(resultsPage.GetRawMatByPage(RAW_MAT_SUPPLIER), String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by supplier FaF"));

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_RawMatByRecipe()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();

            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = TestContext.Properties["Prodman_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            Assert.IsTrue(isServiceNormal, "pas de service Normal ");
            Assert.IsTrue(isServiceValide, "pas de service Valide");

            var resultPage = qtyAjustementPage.GoToResultPage();
           resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.FilterRawBy(RAW_MAT_RECIPE);
            resultPage.UnfoldAll();

            // Récupération de l'association des recettes et de leurs items
            Dictionary<string, List<string>> mapRecipeItems = resultPage.GetRecipeAndItemsAssociated();
            Assert.IsTrue(mapRecipeItems.Count > 0, "no items to map"); 
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
        public void PR_PRODMAN_RawMatByRecipe_FaF()
        {
            // Pour ce test, nous travaillons sur les données de vol pour Jour+1
            // et sur le prod nous travaillons sur le jour J

            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            // On a cree un vol avec customerName1 qui est associer avec RecipeName1
            string RecipeName1 = TestContext.Properties["Prodman_RecipeName1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "Raw mat by recipe type" et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RawMaterialByRecipe, true);
            ResultPage resultsPage = filterAndFavorite.DoneToResults();
            bool isCorrectSelection = resultsPage.IsCorrectOptionSelected();

            // Assert
            Assert.IsTrue(isCorrectSelection, "Raw mat by recipe FaF");


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_RawMatByCustomer()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();

            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();

            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = TestContext.Properties["Prodman_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);

            if (!isServiceNormal || !isServiceValide)
            {
                filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            }

            List<string> customerList = qtyAjustementPage.GetCustomerList();

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_CUSTOMER);
            resultPage.UnfoldAll();

            var listCustomersResults = resultPage.GetCustomerNames();

            bool result = customerList.All(s => listCustomersResults.Contains(s)) && listCustomersResults.All(s => customerList.Contains(s));

            //assert
            Assert.IsTrue(result, "Les customers affichés dans la page Raw mat by customer ne correspondent pas à ceux des services.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_RawMatByCustomer_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par "Raw mat by recipe type" et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RawMaterialByCustomer, true);

            ResultPage resultsPage = filterAndFavorite.DoneToResults();
            // Filtrer pour afficher les données cachées
            resultsPage.FoldAll();

            // Assert 1 : vérifier que la vue des résultats a bien sélectionné "Raw mat by customer"
            Assert.IsTrue(resultsPage.GetRawMatByPage(RAW_MAT_CUSTOMER), String.Format(MessageErreur.FILTRE_ERRONE, "Raw mat by customer FaF"));

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_UnfoldAll()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            bool isServicePresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(isServicePresent, $"Le service {serviceNameNormal} n'est pas visible.");
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            int itemsCount = resultPage.CountItems();
            bool isFoldAll = resultPage.IsFoldAll();
            Assert.AreNotEqual(0, itemsCount, "Il n'y a pas de groupes d'items visibles.");
            Assert.IsTrue(isFoldAll, "Le détail des items visible avant d'avoir activé le UnfoldAll.");
            resultPage.ShowDetail();
            bool isDetailVisible = resultPage.IsDetailVisible();
            Assert.IsTrue(isDetailVisible, "L'option d'unfold via l'item n'est pas fonctionnelle.");
            WebDriver.Navigate().Refresh();
            qtyAjustementPage.PageSize("100");
            resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();
            bool isUnfoldAll = resultPage.IsUnfoldAll();
            Assert.IsTrue(isUnfoldAll, "L'option d'unfoldAll via le menu n'est pas fonctionnelle.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_FoldAll()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            if (!isServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);
                filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            }
            bool isServicePresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(isServicePresent, $"Le service {serviceNameNormal} n'est pas visible.");
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            int itemsCount = resultPage.CountItems();
            bool isFoldAll = resultPage.IsFoldAll();
            Assert.AreNotEqual(0, itemsCount, "Il n'y a pas de groupes d'items visibles.");
            Assert.IsTrue(isFoldAll, "Le détail des items visible avant d'avoir activé le UnfoldAll.");
            resultPage.ShowDetail();
            bool isDetailVisible = resultPage.IsDetailVisible();
            Assert.IsTrue(isDetailVisible, "L'option d'unfold via l'item n'est pas fonctionnelle.");
            resultPage.ShowDetail();
            isDetailVisible = resultPage.IsDetailVisible();
            Assert.IsFalse(isDetailVisible, "L'option de fold via l'item n'est pas fonctionnelle.");
            WebDriver.Navigate().Refresh();
            qtyAjustementPage.PageSize("100");
            resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();
            bool isUnfoldAll = resultPage.IsUnfoldAll();
            Assert.IsTrue(isUnfoldAll, "L'option UnfoldAll via le menu n'est pas fonctionnelle.");
            resultPage.FoldAll();
            bool isFold = resultPage.IsFoldAll();
            Assert.IsTrue(isFold, "L'option FoldAll via le menu n'est pas fonctionnelle.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Show_UseCase()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);

            if (!isServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);

                filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            }

            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            resultPage.UnfoldAll();
            var resultModalPage = resultPage.ShowUseCase();

            //assert
            Assert.IsTrue(resultModalPage.IsServiceVisible(), "La page des use cases n'est pas apparu.");
        }

        //Generate Output form
        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_GenerateOutputForm()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string placeFrom = TestContext.Properties["PlaceFrom"].ToString();
            string placeTo = "Produccion";

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);

            if (!isServiceNormal)
            {
                CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);

                filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
                filterAndFavorite.ResetFilter();
                qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            }

            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            // On récupère la liste des items présents
            List<String> items = resultPage.GetItemNames();

            // Génération de l'outputForm
            string outputFormNumber = resultPage.CreateOutputForm(placeFrom, placeTo);
            var outputFormItemPage = resultPage.GenerateOutputForm();
            outputFormItemPage.Filter(PageObjects.Warehouse.OutputForm.OutputFormItem.FilterItemType.ShowItemsWithPhysQty, true);
            try
            {
                var outformItems = outputFormItemPage.GetItemNames();

                Assert.IsTrue(items.All(s => outformItems.Contains(s)) && outformItems.All(s => items.Contains(s)), "Les items de l'outputform ne sont pas les mêmes que ceux du production management");

                var outputFormPage = outputFormItemPage.BackToList();
                outputFormPage.ResetFilter();
                outputFormPage.Filter(PageObjects.Warehouse.OutputForm.OutputFormPage.FilterType.SearchByNumber, outputFormNumber);

                Assert.AreEqual(outputFormNumber, outputFormPage.GetFirstOutputFormNumber(), "L'output form générée n'est pas dans la liste des outputForms");
            }
            finally
            {
                outputFormItemPage.Close();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_GenerateSupplyOrder()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string deliveryLocation = "Produccion";

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.StartTime, "22:30");
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.EndTime, "23:30");
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(isServiceNormal, "pas de service Normal ");

            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.FilterRawBy(RAW_MAT_GROUP);
            resultPage.UnfoldAll();

            // On récupère la liste des items présents
            List<String> items = resultPage.GetItemNames();

            // Génération supply order
            string supplyOrderNumber = resultPage.CreateSupplyOrder(deliveryLocation);
            SupplyOrderItem supplyOrderItemPage = resultPage.GenerateSupplyOrder();

            try
            {
                List<String> supplyOrderItems = supplyOrderItemPage.GetItemNames();

                Assert.IsTrue(items.All(s => supplyOrderItems.Contains(s)) && supplyOrderItems.All(s => items.Contains(s)), "Les items du supply order ne sont pas les mêmes que ceux du production management");

                SupplyOrderPage supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(SupplyOrderPage.FilterType.ByNumber, supplyOrderNumber);

                Assert.AreEqual(supplyOrderNumber, supplyOrderPage.GetFirstSONumber(), "Le supply order générée n'est pas dans la liste des supply order");
            }
            finally
            {
                supplyOrderItemPage.Close();
            }
        }

        // ________________________________________________ Prints ______________________________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_RawMaterialByWorkshop_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_WORKSHOP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.RawMaterialsByWorkshop, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_RawMaterialByGroup_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.RawMaterialsByGroup, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_Assembly_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.AssemblyReport, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_RecipeReportV2_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.RecipeReportV2, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_RecipeReportDetailedV2_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.RecipeReportDetailedV2, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_HACCPSanitization_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPSanitization, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_HACCPPlating_NewVersion()
        {
            // Pour ce test, nous avons besoin d'un seul service lié à un vol.
            // L'objectif de ce test est de vérifier la fonctionnalité de telechargement de rapport HACCP Plating  .
            // Prepare
            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAdjPage = filterAndFavorite.DoneToQtyAjustement();
            // switch to result Page 
            ResultPage resultsPage =  qtyAdjPage.GoToResultPage(); 
            resultsPage.ClearDownloads();
            PrintReportPage reportPage = resultsPage.PrintReport(ResultPage.PrintType.HACCPPlating, newVersionPrint);
            bool isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert 
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_HACCPPlatingByGroup_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPPlatingGroupByFlight, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            var windowIsOpened = qtyAjustementPage.VerifyIfNewWindowIsOpened();
            Assert.IsTrue(windowIsOpened, "un seul window est ouvert");

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_HACCPPlatingByRecipe_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPPlatingGroupByRecipe, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_HACCPTraySetup_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPTraySetup, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_HACCPHotKitchen_NewVersion()
        {
            // Prepare
            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RawMaterialByRecipe, true);
            ResultPage resultsPage = filterAndFavorite.DoneToResults();
            resultsPage.ClearDownloads();

            PrintReportPage reportPage = resultsPage.PrintReport(ResultPage.PrintType.HACCPHotKitchen, newVersionPrint);
            bool isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_HACCPSlice_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPSlice, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_HACCPThawing_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPThawing, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_DatasheetByRecipe_NewVersion()
        {
            // Prepare
            bool newVersionPrint = true;
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.RawMaterialByRecipe, true);
            ResultPage resultsPage = filterAndFavorite.DoneToResults();
            var reportPage = resultsPage.PrintReport(ResultPage.PrintType.DatasheetByRecipe, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            //assert
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_DatasheetByGuest_NewVersion()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);

            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");

            qtyAjustementPage.ClearDownloads();

            qtyAjustementPage.PageSize("100");
            var resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.FilterRawBy(RAW_MAT_GROUP);

            // Impression
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.DatasheetByGuest, newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Favoris_Delete_FaF()
        {
            // Le test vérifie le bon fonctionnement de la suppression d'un favori créé.
            //Prepare 
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string favouriteName = RandomString(5);

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.MakeFavorite(favouriteName);

            filterAndFavorite.SetFavoriteText(favouriteName);
            Assert.IsTrue(filterAndFavorite.IsFavoritePresent(favouriteName), "Le favori " + favouriteName + " n'a pas été créé.");
            filterAndFavorite.DeleteFavorite(favouriteName);

            //Assert
            Assert.IsFalse(filterAndFavorite.IsFavoritePresent(favouriteName), "Le favori " + favouriteName + " n'a pas été supprimé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Favoris_Goto_FaF()
        {
            // On n'a pas besoin de données pour ce test, on va juste vérifier l'application des filtres de l'interface
            // Filtre et favoris vers l'interface de gestion de production

            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string favouriteName = RandomString(5);
            DateTime dtfrom = DateTime.Today.AddDays(-11);
            DateTime dtto = DateTime.Today;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Appliquer les filtres
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateFrom, dtfrom);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateTo, dtto);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Customers, customerName1);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.GuestType, guestType1);
            filterAndFavorite.MakeFavorite(favouriteName);
            filterAndFavorite.SetFavoriteText(favouriteName);
            Assert.IsTrue(filterAndFavorite.IsFavoritePresent(favouriteName), "Le favori " + favouriteName + " n'a pas été créé.");
            ResultPage resultPage = filterAndFavorite.SelectFavoriteName(favouriteName);

            // Vérification des filtres sur l'interface des résultats
            DateTime datefrom = resultPage.getDateFilterFrom();
            DateTime dateto = resultPage.getDateFilterTo();
            string siteFltr = resultPage.GetFilterValue("SiteId");
            List<string> seledFltrCust = resultPage.GetSelectedCustomers();
            List<string> seledFltrGuestTyp = resultPage.GetSelectedGuestType();
            bool rsltDatefrom = DateTime.Equals(datefrom, dtfrom);
            bool rsltDateto = DateTime.Equals(dateto, dtto);

            // Assert
            Assert.IsTrue(rsltDatefrom, "Le filtre DateFrom n'a pas été appliqué.");
            Assert.IsTrue(rsltDateto, "Le filtre DateTo n'a pas été appliqué.");
            Assert.AreEqual(siteACE, siteFltr, "Le filtre site n'a pas été appliqué.");
            Assert.IsTrue(seledFltrCust.Any(s => s.Contains(customerName1)));
            Assert.IsTrue(seledFltrGuestTyp.Any(s => s.Contains(guestType1)));

            filterAndFavorite = resultPage.Back();
            filterAndFavorite.SetFavoriteText(favouriteName);
            filterAndFavorite.DeleteFavorite(favouriteName);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Favoris_Rename_FaF()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string favouriteName = RandomString(5);
            string newFavouriteName = RandomString(5);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Création du favori
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.MakeFavorite(favouriteName);
            filterAndFavorite.SetFavoriteText(favouriteName);
            Assert.IsTrue(filterAndFavorite.IsFavoritePresent(favouriteName), "Le favori " + favouriteName + " n'a pas été créé.");

            // Modification du nom
            EditFavoriteModal editFavoriteModal = filterAndFavorite.EditFavoriteName(favouriteName);
            editFavoriteModal.RenameFavorite(newFavouriteName);
            editFavoriteModal.SaveEdit();
            filterAndFavorite.SetFavoriteText(newFavouriteName);
            string nouveauNomFavRecupere = filterAndFavorite.GetFavoriteNameText();

            // Assert
            Assert.IsTrue(newFavouriteName.Equals(nouveauNomFavRecupere), "Le nom du filtre favori n'a pas été renommé.");

            filterAndFavorite.DeleteFavorite(newFavouriteName);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Favoris_Search_FaF()
        {
            // On n'a pas besoin de données côté vol et service pour effectuer ce test

            //Prepare 
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string favouriteName = RandomString(5);

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);

            // Création du favori
            filterAndFavorite.MakeFavorite(favouriteName);
            filterAndFavorite.SetFavoriteText(favouriteName);

            //Assert
            Assert.IsTrue(filterAndFavorite.IsFavoritePresent(favouriteName), "Le favori " + favouriteName + " n'a pas été créé.");

            // Suppression de favorite deja cree 
            filterAndFavorite.DeleteFavorite(favouriteName);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_RestFilter_FaF()
        {
            // Prepare
            //string siteIBZ = TestContext.Properties["Site"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            DateTime dtfrom = DateTime.Today.AddDays(-11);
            DateTime dtto = DateTime.Today.AddDays(2);

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            //filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteIBZ);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateFrom, dtfrom);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateTo, dtto);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Customers, customerName1);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.GuestType, guestType1);

            var listGestTypBeforeReset = filterAndFavorite.GetSelectedFilters("SelectedGuestTypesTrueValue");
            var listcustBeforeReset = filterAndFavorite.GetSelectedFilters("SelectedGuestTypesTrueValue");
            var datefromBeforeReset = filterAndFavorite.getDateFilterFrom();
            var datetoBeforeReset = filterAndFavorite.getDateFilterTo();
            var siteFltrBeforeReset = filterAndFavorite.GetFilterValue("SiteId");
            filterAndFavorite.ResetFilter();

            var listGestTypAfterReset = filterAndFavorite.GetSelectedFilters("SelectedGuestTypesTrueValue");
            var listcustAfterReset = filterAndFavorite.GetSelectedFilters("SelectedGuestTypesTrueValue");
            var datefromAfterReset = filterAndFavorite.getDateFilterFrom();
            var datetoAfterReset = filterAndFavorite.getDateFilterTo();
            //var siteFltrAfterReset = filterAndFavorite.GetFilterValue("SiteId");


            //Assert
            Assert.AreNotEqual(listGestTypBeforeReset, listGestTypAfterReset);
            Assert.AreNotEqual(listcustBeforeReset, listcustAfterReset);
            Assert.AreNotEqual(datefromBeforeReset, datefromAfterReset);
            Assert.AreNotEqual(datetoBeforeReset, datetoAfterReset);
            //Assert.AreNotEqual(siteFltrBeforeReset, siteFltrAfterReset);

            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, "MAD");
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, "ACE");
        }

        [Ignore] //FoodPacket généré pas service externe. Customers>Food Packet
        [TestMethod]
        public void PR_PRODMAN_Filter_FoodPacket()
        {
            //Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("yyyy-MM-dd");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            //Avoir des données disponibles.
            //Pour cela il faut qu'un vol (onglet "flight") soit validé et lié au datasheet.
            CreateFlight(homePage, siteACE, flightNameNormal, customerName1, guestType1, serviceNameNormal);
            //Pour lier un vol à une datasheet
            //il faut aller dans service/ price,
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameNormal);
            ServicePricePage service = servicePage.ClickOnFirstService();
            //cliquer sur le symbole "stylo"
            //puis entrer la datasheet à lier.
            service.UnfoldAll();
            string datasheetName = service.GetDatasheet();
            Assert.AreEqual("datasheetProdman1", datasheetName);
            // datasheet>Repice>Item
            DatasheetPage datasheetIndex = homePage.GoToMenus_Datasheet();
            datasheetIndex.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            DatasheetDetailsPage datasheet = datasheetIndex.SelectFirstDatasheet();
            Assert.AreEqual("recetteProdman1", datasheet.GetFirstRecipeName());
            RecipesPage recipesIndex = homePage.GoToMenus_Recipes();
            recipesIndex.Filter(RecipesPage.FilterType.SearchRecipe, "recetteProdman1");
            RecipeGeneralInformationPage recipe = recipesIndex.SelectFirstRecipe();
            Assert.AreEqual("recetteProdman1", recipe.GetRecipeName());


            //1. Appliquer les filtres sur food packet,
            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            // FIXME : nourrir le combobox FoodPacket
            prodManagement.Filter(FilterAndFavoritesPage.FilterType.FoodPacket, "None");
            //Vérifier que les résultats s'accordent bien au filtre appliqué
            QuantityAdjustmentsPage adj = prodManagement.DoneToQtyAjustement();
            ResultPage result = adj.GoToResultPage();

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_Customer_QA()
        {
            // Prepare
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            // Get Customer ICAO for the Assert 2 
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customerName1);
            string customerIcao = customersPage.GetFirstCustomerIcao();
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // Filtre par workshop et par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Customers, customerName1);
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            string selectedCustomer = qtyAjustementPage.GetCustomer();
            List<string> selectedCustomerList = qtyAjustementPage.GetCustomerList();
            foreach (var item in selectedCustomerList)
            {
                // Assert 1 : The customer is well selected 
                Assert.IsTrue(selectedCustomer.Contains(item), String.Format(MessageErreur.FILTRE_ERRONE, "Customer"));
            }
            string customerICAOfromQty = qtyAjustementPage.GetFirstServiceICAO();
            //Assert 2 : The service shown is for customer1 
            Assert.AreEqual(customerIcao, customerICAOfromQty, "le filtre de client n est pas pris en charge .");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_GenerateSO_Result()
        {
            string ChooseOneOption = "Produccion";

            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            // "Select a place…" non sélectionnable
            string supplyOrderNoOptionModal = resultPage.CreateSupplyOrder(null);
            SupplyOrderItem supplyOrderNoOption = resultPage.GenerateSupplyOrder(true);


            string delieveryIsRequired = resultPage.GetDelieveryIsRequiredMessage();
            Assert.AreEqual("The delivery place is required.", delieveryIsRequired);

            resultPage.ChooseDeliveryLocationOption(ChooseOneOption);

            SupplyOrderItem supplyOrderWithOption = resultPage.GenerateSupplyOrder();
            string supply_order_NO = supplyOrderWithOption.GetSupplyOrderFormNumber();
            Assert.AreEqual(supply_order_NO, supplyOrderNoOptionModal, "The supply order NO is not the same between the form created and the supply order page ");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ShowVaccumOnly_QA()
        {
            // Pour ce test, nous avons besoin de deux services : normal et vacuum.
            // L'objectif de ce test est de tester la fonctionnalité du filtre service sur le service vacuum. C'est pour cela que nous avons créé 2 services.
            // Cela nous permettra d'obtenir le résultat du filtre sur vacuum.

            // Prepare 
            string serviceNameVacuum = TestContext.Properties["Prodman_ServiceVacuum"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceVavccum = qtyAjustementPage.IsServicePresent(serviceNameVacuum);
            Assert.IsTrue(isServiceNormal, "le service nnormal nest pas existant");
            Assert.IsTrue(isServiceVavccum, "le service flight valide n'est pas existant");
            // filtrer Vaccum
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowVacuumProd, true);
            string serivce_name_affichee = qtyAjustementPage.GetFirstServiceName();
            //Assert 1 : Verifier le nom de service Vaccum 
            Assert.AreEqual(serivce_name_affichee, serviceNameVacuum, "Le filtrage par ShowVaccumOnly n'a pas été appliqué.");
         }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ServiceType_QA()
        {
            string serviceCategory1 = TestContext.Properties["Prodman_Needs_ServiceCategory1"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            var resultBeforeFilter = qtyAjustementPage.CountServices();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ServicesCategorie, serviceCategory1);
            var resultAfterFilter = qtyAjustementPage.CountServices();
            Assert.AreNotEqual(resultBeforeFilter, resultAfterFilter, "Le filtrage par Service Categories  n'a pas été appliqué.");
            ServicePricePage servicePricePage = qtyAjustementPage.EditFirstService();
            ServiceGeneralInformationPage serviceGeneralInformationPage = servicePricePage.ClickOnGeneralInformationTabOtherNavigation();
            var gategoryService = serviceGeneralInformationPage.GetCategory();
            Assert.AreEqual(gategoryService, serviceCategory1, "Le filtrage par Service Categories  n'a pas été appliqué.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_Service_QA()
        {
            // Pour que le service soit visible sur l'interface QA,
            // il est nécessaire de créer un vol J+1, sinon il ne sera pas disponible, même pour le filtre.
            // L'objectif de ce test :  est de vérifier la prise en compte du filtre de service SUR L'INTERFACE DE QA (PROD MAN).
            // !!! LA VÉRIFICATION DES DONNÉES AFFICHÉES CONCERNE LE TEST 'PR_PRODMAN_FILTER_SERVICE' !!! 

            string serviceProdmanNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();

            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceProdmanNormal);
            // Verifier le filtre a ete pris en compte 
            bool is_service_available = qtyAjustementPage.IsFilterServiceApplied(serviceProdmanNormal); 
            Assert.IsTrue(is_service_available, "le filtre n'a pas été pris en compte , verifier la creation d un vol avec un service normal ");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_StartEndTime_Result()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();

            // Filtre par StartTime & EndTime
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            var serviceCountWithFirstFilter = qtyAjustementPage.CountServices();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.StartTime, "23:58");
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.EndTime, "23:59");
            var serviceCountWithSecoundFilter = qtyAjustementPage.CountServices();

            //assert
            Assert.IsTrue(!serviceCountWithFirstFilter.Equals(serviceCountWithSecoundFilter), String.Format(MessageErreur.FILTRE_ERRONE, "StartTime & EndTime"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_FromTo_QA()
        {
            DateTime dateFrom = DateUtils.Now.AddDays(-7);
            DateTime dateTo = DateUtils.Now.AddDays(+1);
            DateTime dateFromQuantityAdjustmentsPage = DateUtils.Now.AddDays(-1);
            DateTime dateToQuantityAdjustmentsPage = DateUtils.Now.AddDays(+7);
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateFrom, dateFrom);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateTo, dateTo);
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            int serviceCountWithFirstFilter = qtyAjustementPage.CountServices();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateFrom, dateFromQuantityAdjustmentsPage);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, dateToQuantityAdjustmentsPage);
            List<string> servicesDisplayed = qtyAjustementPage.GetServicesName();
            foreach (string service in servicesDisplayed)
            {
                ServicePricePage servicePricePage = qtyAjustementPage.EditService(service);
                servicePricePage.Go_To_New_Navigate();
                servicePricePage.UnfoldAll();
                ServiceCreatePriceModalPage servicePriceModal = servicePricePage.EditFirstPriceService();
                string serviceDateFrom = servicePriceModal.GetServicePeriodFrom();
                string serviceDateTo = servicePriceModal.GetServicePeriodTo();
                DateTime datePriceServiceFrom = DateTime.ParseExact(serviceDateFrom, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                DateTime datePriceServiceTo = DateTime.ParseExact(serviceDateTo, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                servicePriceModal.Close();
                bool isServiceStartDateValid = datePriceServiceFrom <= dateFromQuantityAdjustmentsPage;
                bool isServiceEndDateValid = datePriceServiceTo >= dateToQuantityAdjustmentsPage;
                Assert.IsTrue(isServiceStartDateValid && isServiceEndDateValid, string.Format(MessageErreur.FILTRE_ERRONE, "dateFrom & dateTo"));
                servicePricePage.Close();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_GenerateOF_Result()
        {
            string placeFromWithOption = "Produccion";
            string placeToWithOption = "Economato";
            string comment = "this is a comment";
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            // "SelectPlace" non sélectionnable
            string outputFormNoOptionModal = resultPage.CreateOutputForm(null, null);
            OutputFormItem outfputFormNoOption = resultPage.GenerateOutputForm(true);

            string fromPlaceIsRequired = resultPage.GetFromPlaceIsRequiredMessage();
            Assert.AreEqual("The place 'from' is required.", fromPlaceIsRequired);

            string toPlaceIsRequired = resultPage.GetToPlaceIsRequiredMessage();
            Assert.AreEqual("The place 'to' is required.", toPlaceIsRequired);
            // Fill the form 
            resultPage.ChooseFromAndToPlaceOptions(placeToWithOption, placeFromWithOption, comment);
            // Verifty the OF number 
            OutputFormItem outputFormWithOption = resultPage.GenerateOutputForm();
            string of_number = outputFormWithOption.GetOutputFormNumber();
            Assert.AreEqual(of_number, outputFormNoOptionModal, "the OF number from the form is not the same as the OF number in the Output Form ");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_ServiceGlobalView3_QA()
        {
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            //Arrange 
            HomePage homePage = LogInAsAdmin();

            string decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            double initQty = qtyAjustementPage.GetQuantity(decimalSeparatorValue, serviceNameNormal);

            qtyAjustementPage.ConsultService(serviceNameNormal);
            qtyAjustementPage.SetQuantityAfterConsult((initQty + 1).ToString());
            //Assert1 : check the value 
            Assert.IsTrue(qtyAjustementPage.IsQuantityUpdatedAfterConsult((initQty + 1).ToString()), "Modification failed");
            // assert 2 : chagement en bleu 
            bool changement_de_couleur = qtyAjustementPage.IsQtyElementAdjusted();
            Assert.IsTrue(changement_de_couleur, "apres la modification de qty le couleur de champ doit etre bleu  ");
            qtyAjustementPage.ReturnTothePastQty();
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_ServiceGlobalView2_QA()
        {
            string servicePromanNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            qtyAjustementPage.SelectService(servicePromanNormal);
            qtyAjustementPage.OpenNewQtyDetailsPopUp();
            Assert.IsTrue(qtyAjustementPage.AreAllFlightsDifferent(), "le filtre des détails Quantity Adjustment dans la PopUp n'a pas été pris en compte");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_ServiceGlobalView4_QA()
        {
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNormal = TestContext.Properties["Prodman_FlightNormal"].ToString() + DateTime.Today.AddDays(1).ToString("dd/MM/yyyy"); 
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.SetDateState(DateUtils.Now.AddDays(1));
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNormal);
            FlightDetailsPage editPage =  flightPage.EditFirstFlight(flightNormal); 
            string flightServiceQuantity = editPage.GetServiceQty ();
            string decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.ConsultService(serviceNameNormal);
            string  qtyValueAfterRefresh = qtyAjustementPage.ClickRefreshAndGetTheQty();

            Assert.AreEqual(flightServiceQuantity, qtyValueAfterRefresh);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ItemSubgroup_Result()
        {
            string NORWEGIANItem = TestContext.Properties["Prodman_Needs_Item_Group4"].ToString(); // NORWEGIAN
            string itemGroup = "ItemGroup";
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowHiddenArticles, true);
            int resLengthBeforeFilter = resultPage.CountResults();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ItemGroups, NORWEGIANItem);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowHiddenArticles, true);
            int resLengthAfterFilter = resultPage.CountResults();
            string itemSelected = qtyAjustementPage.GetSelectedGroupBy();
            string resultFilter = resultPage.GetFirstItemGroup();

            if (resLengthAfterFilter < resLengthBeforeFilter && itemSelected == itemGroup)
            {
                Assert.AreEqual(NORWEGIANItem, resultFilter, "Le premier groupe d'articles des résultats devrait être égal à AirCanada.");
            }
            else
            {
                Assert.Fail("Le filtre item Subgroup n'a pas été pris en compte correctement.");
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_QuantityAdjustment_FaF()
        {
            string tabquantityAdjustmentsName = "Quantity adjustments";
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            prodManagement.Filter(FilterAndFavoritesPage.FilterType.QtyAdjustements, true);
            ResultPage qtyAjustementTab = prodManagement.DoneToResults();
            string activatedNavTab = qtyAjustementTab.GetActivatedNavTab();
            Assert.AreEqual(tabquantityAdjustmentsName.ToLower(), activatedNavTab.ToLower(), "Vous êtes hors de la vue Quantity adjustment");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ShowNormalOnly_QA()
        {
            // Pour ce test, nous avons besoin de deux services : normal et vacuum.
            // L'objectif de ce test est de tester la fonctionnalité du filtre service sur le service Normal. C'est pour cela que nous avons créé 2 services.
            // Cela nous permettra d'obtenir le résultat du filtre sur Normal .

            // Prepare 
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowNormalProd, true);
            string serivce_name_affichee = qtyAjustementPage.GetServiceName();
            Assert.AreEqual(serivce_name_affichee, serviceNameNormal , "Le filtrage par ShowNormalProd n'a pas été appliqué.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ShowNormalVaccum_QA()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();


            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowNormalProd, true);
            var resultFilterShowNormalProductionOnly = qtyAjustementPage.CountServices();

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowVacuumProd, true);
            var resultFilterShowVacuumProductionOnly = qtyAjustementPage.CountServices();

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowNormalAndVacuumProd, true);
            var resultFilterShowNormalAndVacuumProd = qtyAjustementPage.CountServices();
            Assert.AreEqual(resultFilterShowNormalProductionOnly + resultFilterShowVacuumProductionOnly, resultFilterShowNormalAndVacuumProd, "Le filtrage par Show normal et vacuum production n'a pas été appliqué.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_StartEndTime_FaF()
        {
            // Prepare
            string startTimeFaF = "12:30";
            string endTimeFaF = "20:00";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // On filtre les données par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.StartTime, startTimeFaF);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.EndTime, endTimeFaF);
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            var startTime = qtyAjustementPage.GetStartTime();
            var endTime = qtyAjustementPage.GetEndTime();
            Assert.AreEqual(startTimeFaF, startTime, "Le filtre par start time n'a pas été appliqué.");
            Assert.AreEqual(endTimeFaF, endTime, "Le filtre par end time n'a pas été appliqué.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ItemSubgroup_FaF()
        {
            // Prepare
            string subgroupeFaf = "subgrpname";
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            bool subGroupAvant = homePage.GetSubGroupFunctionValue();
            homePage.SetSubGroupFunctionValue(true);

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            // On filtre les données par site
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.ItemSubGroups, subgroupeFaf);
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            string subgroupQA = qtyAjustementPage.GetItemSubGroup();

            // Assert 1 :Verifier que le filtre est appliquer dans section de filtre 
            Assert.AreEqual(subgroupeFaf, subgroupQA, "Le filtre par SubGroupe n'a pas été appliqué.");
            homePage.SetSubGroupFunctionValue(subGroupAvant);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_OpenService_QA()
        {
            // Pour ce test, nous avons besoin d'un seul service lié à un vol.
            // L'objectif de ce test est de vérifier la fonctionnalité du crayon à côté du service sur l'interface QA .

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();

            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            // Click sur le crayon 
            ServicePricePage servicePricePage = qtyAjustementPage.EditFirstService();

            Assert.IsTrue(servicePricePage.CheckActiveService(), "Une nouvelle fenêtre s’ouvre sur le service ");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_UseCaseCRecette_Result()
        {
            // Pour ce test, nous avons besoin d'un seul service.
            // L'objectif de ce test est de tester la fonctionnalité du crayon de service d'après l'interface Results (Prod Man).

            // Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
          
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.UnfoldAll();
            string  item = resultPage.GetFirstItemName();
            ResultModal resultModal = resultPage.GoToUseCaseModalPage(item, 0);
            resultModal.UnfoldAll();
            resultModal.FirstEditRecipe();
            //assert
            Assert.IsTrue(resultModal.VerifyRecipePageOpened(), "La page Recipe ne s'ouvre pas.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_TraySetUp_Result()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Raw materials Report_-_";
            string DocFileNameZipBegin = "All_files_";
            //Arrange
            HomePage homePage = LogInAsAdmin();
            homePage.Navigate();
            homePage.ClearDownloads();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            var resultPage = qtyAjustementPage.GoToResultPage();
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPTraySetup, true);
            var isreportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            Assert.IsTrue(isreportGenerated, "Le document PDF n'a pas pu être généré par l'application.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_TraySetUpMultiflight_Result()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Raw materials Report_-_";
            string DocFileNameZipBegin = "All_files_";
            //Arrange
            HomePage homePage = LogInAsAdmin();

            homePage.ClearDownloads();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            var resultPage = qtyAjustementPage.GoToResultPage();
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.Haccptraysetupmultiflight, true);
            var isreportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            Assert.IsTrue(isreportGenerated, "Le document PDF n'a pas pu être généré par l'application.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_TraySetUpMultiflightnoguest_Result()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Raw materials Report_-_";
            string DocFileNameZipBegin = "All_files_";
            //Arrange
            HomePage homePage = LogInAsAdmin();

            homePage.ClearDownloads();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            var resultPage = qtyAjustementPage.GoToResultPage();
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.Haccptraysetupmultiflightnoguest, true);
            var isreportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            Assert.IsTrue(isreportGenerated, "Le document PDF n'a pas pu être généré par l'application.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_UseCaseCalcul_Result()
        {
            string airCanadaItem = TestContext.Properties["Prodman_Needs_Item_Group2"].ToString();
            string service = "service";
            DateTime dateService;
            double numberqte;
            double number = 0.0;
            bool isValidNumberqte, isValidNumber = false;
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            int resLengthBeforeFilter = resultPage.CountResults();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ItemGroups, airCanadaItem);
            int resLengthAfterFilter = resultPage.CountResults();
            string resultFilter = resultPage.GetFirstItemGroup();
            if ((airCanadaItem == resultFilter && resLengthAfterFilter > 0) || resLengthBeforeFilter > 0)
            {
                resultPage.UnfoldAll();
                string item = resultPage.GetFirstItemName();
                string itemQte = resultPage.GetFirstItemQty();
                ResultModal resultModal = resultPage.GoToUseCaseModalPage(item, 0);
                resultModal.UnfoldAll();
                bool displayModelServiceUseCase = resultModal.IsServiceVisible();
                if (displayModelServiceUseCase)
                {
                    string servicename = resultModal.GetServiceName_UseCase();
                    string date_usecase = resultModal.GetDate_UseCase();
                    string flight_usecase = resultModal.GetFlight_UseCase();
                    string date_detail_usecase = resultModal.GetDate_Detail_UseCase();
                    string customer_usecase = resultModal.GetCustomer_UseCase();
                    string service_usecase = resultModal.GetService_UseCase();
                    string service_qte_usecase = resultModal.GetSERVICE_QTE_UseCase();
                    string recipe_usecase = resultModal.GetRecipe_UseCase();
                    string nb_occ_usecase = resultModal.GetNB_OCC_UseCase();
                    string qte_usecase = resultModal.GetQTEItem_UseCase();
                    string yield_usecase = resultModal.GetYield_UseCase();
                    string methode_usecase = resultModal.GetMethode_UseCase();
                    string Coefficient_usecase = resultModal.GetCoefficient_UseCase();
                    string quantity_usecase = resultModal.GetQuantity_UseCase();
                    bool startsWith_Service = servicename.StartsWith(service, StringComparison.OrdinalIgnoreCase);
                    string[] formats = { "dd/MM/yyyy", "d/M/yyyy" };
                    bool isValidDate = DateTime.TryParseExact(date_usecase, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateService);
                    string[] parts = itemQte.Split(' ');
                    if (parts.Length > 0)
                        isValidNumber = double.TryParse(parts[0], out number);
                    Assert.IsTrue(startsWith_Service, "Le Modal ne commence pas par un service.");
                    Assert.IsTrue(isValidDate, "La date est incorrecte.");
                    Assert.IsTrue(!string.IsNullOrEmpty(flight_usecase), "Le flight est vide.");
                    Assert.IsTrue(!string.IsNullOrEmpty(date_detail_usecase), "Le date est vide.");
                    Assert.IsTrue(!string.IsNullOrEmpty(customer_usecase), "Le customer est vide.");
                    Assert.IsTrue(!string.IsNullOrEmpty(service_usecase), "Le service est vide.");
                    Assert.IsTrue(!string.IsNullOrEmpty(recipe_usecase), "Le recipe est vide.");
                    Assert.IsTrue(!string.IsNullOrEmpty(service_qte_usecase), "Le service quantity est vide.");
                    Assert.IsTrue(!string.IsNullOrEmpty(nb_occ_usecase), "Le nb occ est vide.");
                    Assert.IsTrue(!string.IsNullOrEmpty(qte_usecase), "Le quantity est vide.");
                    Assert.IsTrue(!string.IsNullOrEmpty(yield_usecase), "Le yield est vide.");
                    Assert.IsTrue(!string.IsNullOrEmpty(methode_usecase), "La methode est vide.");
                    Assert.IsTrue(!string.IsNullOrEmpty(Coefficient_usecase), "Le coefficient est vide.");
                    Assert.IsTrue(!string.IsNullOrEmpty(quantity_usecase), "Le quantity est vide.");
                    isValidNumberqte = double.TryParse(quantity_usecase, out numberqte);
                    if (isValidNumberqte && isValidNumber)
                        Assert.IsTrue(numberqte == number, "La quantité n'est pas la meme que dans la page principale.");
                }
                resultModal.CloseModal();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_ServiceGlobalView1_QA()
        {
            // Pour ce test, nous avons besoin de deux services .
            // L'objectif de ce test est de vérifier la fonctionnalité de fleche suivant sur l'interface qty detqils popUP sue QA .

            string servicePromanNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            qtyAjustementPage.SelectService(servicePromanNormal);
            qtyAjustementPage.OpenNewQtyDetailsPopUp();
            string previousFlight = qtyAjustementPage.GetFlightFromPopUp();
            // verifier la visibiliter de fleche suivant 
            bool isNextAvailable = qtyAjustementPage.IsNextAvailable();
            Assert.IsTrue(isNextAvailable, "Verifier qu il existe deux service associer a deux vol ");
            qtyAjustementPage.ClickOnNext();
            string nextFlight = qtyAjustementPage.GetFirstFlightNo();
            // les deux vol ne doive pas avoir le meme nom pour assuer la fonctionnaliter du button NEXT 
            Assert.IsFalse(nextFlight.Equals(previousFlight), "Les vols ne sont pas le même");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ItemGroup_woSubgroup_QA()
        {
            string groupfiltreitem = TestContext.Properties["Prodman_Needs_Item_Group2"].ToString().ToUpper().Trim();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ItemGroups, groupfiltreitem);
            var resultPage = qtyAjustementPage.GoToResultPage();
            var resLengthAfterFilter = resultPage.CountResults();
            var resultFilter = resultPage.GetFirstItemGroup();
            if ((groupfiltreitem == resultFilter && resLengthAfterFilter > 0))
            {
                resultPage.UnfoldAll();
                var itemresult = resultPage.GetFirstItemName();
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemresult);
                if (itemPage.CheckTotalNumber() > 0)
                {
                    ItemGeneralInformationPage itemGeneralInformationPage = itemPage.ClickOnFirstItem();
                    var libelle_group = itemGeneralInformationPage.GetGroupName().ToUpper().Trim();

                    Assert.IsTrue(libelle_group == groupfiltreitem, " le filtre ne pris pas en compte : le service qui est lié à la datasheet qui contient cet item.");

                }
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ItemSubgroup_wSubgroup_QA()
        {
            string groupfiltreitem = TestContext.Properties["Prodman_Needs_Item_Group2"].ToString();
            string groupfiltresupitem = "subgrpname";
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();

            //Arrange 
            HomePage homePage = LogInAsAdmin();

            bool subGroupAvant = homePage.GetSubGroupFunctionValue();
            homePage.SetSubGroupFunctionValue(true);

            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();

            // Verify the service exists 
            bool IsServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(IsServiceNormal, "Service pas disponible , verifier qu il y a un vol cree avec un service normal ");
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ItemGroups, groupfiltreitem);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ItemSubGroups, groupfiltresupitem);
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            int resLengthAfterFilter = resultPage.CountResults();

            // Assert  : l existance d item 
            Assert.IsTrue(resLengthAfterFilter >= 1, "No group shown , check the item supgroup in the item section ");
            var resultFilter = resultPage.GetFirstItemGroup();
            if ((groupfiltreitem == resultFilter && resLengthAfterFilter > 0))
            {
                resultPage.UnfoldAll();
                string itemresult = resultPage.GetFirstItemName();
                ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemresult);
                if (itemPage.CheckTotalNumber() > 0)
                {
                    ItemGeneralInformationPage itemGeneralInformationPage = itemPage.ClickOnFirstItem();
                    string libelle_sup_group = itemGeneralInformationPage.getSubGroupeName();
                    Assert.IsTrue(libelle_sup_group == groupfiltresupitem, " le filtre n est pas pris en compte : le service qui est lié à la datasheet qui contient cet item Subgroup.");
                }
            }
            else
            {
                Assert.IsTrue(resLengthAfterFilter < 0 , "No service available "); 
            }

            homePage.SetSubGroupFunctionValue(subGroupAvant);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Print_SliceHotKitchen_Result()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "HACCP SliceHotkitchen_-_";
            string DocFileNameZipBegin = "All_files_";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            var resultPage = qtyAjustementPage.GoToResultPage();
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPSlicekitchen, true);
            reportPage.Close();
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_NoCancelOrInvoiceFlight()
        {
            // Pour ce test, nous avons besoin d'un seul service lié à un vol , on debut le le vol sera prevalidated , puis Invoiced puis cancelled .
            // L'objectif de ce test est de verifier que les services des vols invoiced et cancelled ne serons pas affichers sur QA . 

            string previousFlight = "";
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();

            // Arrange 
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            string serviceName = qtyAjustementPage.GetFirstServiceName();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceName);
            bool IsServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(IsServiceNormal, "Service pas disponible , verifier qu il y a un vol cree avec un service normal ");
            try
            {
                qtyAjustementPage.SelectService(serviceName);
                qtyAjustementPage.OpenNewQtyDetailsPopUp();
                previousFlight = qtyAjustementPage.GetFirstFlightNo();

                // On commance par le vol Invoiced 
                FlightPage flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, previousFlight);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.SetNewState("I");

                // Retour au Prod man 
                prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
                prodManagement.ResetFilter();
                qtyAjustementPage = prodManagement.DoneToQtyAjustement();

                // Verifier que le service apres etre mis invoiced n est plus disponible 
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceName);
                bool After_IsServiceNormal = qtyAjustementPage.IsServicePresent(serviceName);
                Assert.IsFalse(qtyAjustementPage.IsServicePresent(serviceName), "Le service apres l annulation de vol est encore disponible .");

                // La verification de vol cancelled 
                // changement de l'etat de vol vers cancelled 
                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, previousFlight);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.SetNewState("C");

                // Retour au Prod man 
                prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
                prodManagement.ResetFilter();
                qtyAjustementPage = prodManagement.DoneToQtyAjustement();

                // Verifier que le service apres Invoice n est plus disponible 
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceName);
                After_IsServiceNormal = qtyAjustementPage.IsServicePresent(serviceName);
                Assert.IsFalse(qtyAjustementPage.IsServicePresent(serviceName), "Le service apres l annulation de vol est encore disponible .");

            }
            finally
            {
                FlightPage flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, previousFlight);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.WaiForLoad();
                flightPage.UnSetNewState("C");
                WebDriver.Navigate().Refresh();
                flightPage.SetNewState("P");
                Assert.IsTrue(!qtyAjustementPage.IsServicePresent(serviceName), "Le service apres l activation de vol est non disponible ");
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_GuestType_QA()
        {
            // Pour ce test, nous avons besoin d'un service NORMAL avec GUESTTYPE1 = "FC" --> création d'un vol J+1.
            // L'objectif de ce test est de vérifier la fonctionnalité du filtre "guest type" sur l'interface QA.

            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string flightNameNormal = "" + TestContext.Properties["Prodman_FlightNormal"].ToString() + DateTime.Today.AddDays(1).ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.GuestType, guestType1);

            string serviceName = qtyAjustementPage.GetServiceName();
            qtyAjustementPage.DisplayRow(1);
            qtyAjustementPage.OpenNewQtyDetailsPopUp();
            string flightNameFromeQA = qtyAjustementPage.GetFlightFromPopUp();
          
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameFromeQA);
            string preval = "Preval";
            flightPage.Filter(FlightPage.FilterType.Status, preval);
            flightPage.SetDateState(DateUtils.Now.AddDays(1));
            FlightDetailsPage flightDetails = flightPage.EditFirstFlight(flightNameFromeQA);
            string guestType = flightDetails.GetGuestType();
            string flightServiceName = flightDetails.GetServiceName();
            Assert.AreEqual(flightServiceName, serviceName, $"Les vols {flightServiceName} et {serviceName} ne sont pas le même ");
            Assert.AreEqual(guestType1, guestType, $"Les GuestTypes {guestType1} et {guestType} ne sont pas le même");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_CheckResultPage_QA()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();
            string serviceNamedefault = "";
            string datasheetNameDefault = "";
            try
            {
                FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
                prodManagement.ResetFilter();
                QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
                serviceNamedefault = qtyAjustementPage.GetFirstServiceName();
                ServicePricePage servicePricePage = qtyAjustementPage.EditFirstService();
                servicePricePage.Go_To_New_Navigate();
                servicePricePage.UnfoldAll();
                datasheetNameDefault = servicePricePage.GetDatasheet();
                ServiceGeneralInformationPage serviceGeneralInformationPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationPage.SetNotProduced();
                serviceGeneralInformationPage.Close();

                prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
                prodManagement.ResetFilter();
                qtyAjustementPage = prodManagement.DoneToQtyAjustement();
                Assert.IsFalse(qtyAjustementPage.IsServicePresent(serviceNamedefault));
            }
            finally
            {
                ServicePage servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceNamedefault);
                ServicePricePage servicePricePageCustomer = servicePage.ClickOnFirstService();
                ServiceGeneralInformationPage serviceGeneralInformationPageCustomer = servicePricePageCustomer.ClickOnGeneralInformationTab();
                serviceGeneralInformationPageCustomer.SetProduced(true);
                ServicePricePage servicePricePage = serviceGeneralInformationPageCustomer.GoToPricePage();
                servicePricePage.UnfoldAll();
                servicePricePage.SetFirstPriceDataSheet(datasheetNameDefault);
                FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
                QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNamedefault);
                Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNamedefault));
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_RecipeType_QA()
        {
            // Pour ce test, nous avons besoin d'un service name avec un datasheet qui a une recette de type = ENSALADA (recipeType1).
            // L'objectif de ce test est de vérifier la fonctionnalité du filtre recipe type, puis la vérification du type de recette d'après le datasheet.

            string recipeType1 = TestContext.Properties["Prodman_Needs_RecipeType2"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
        
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.RecipeType, recipeType1);
            ServicePricePage editFirstService = qtyAjustementPage.EditFirstService();
            editFirstService.Go_To_New_Navigate();
            editFirstService.UnfoldAll();
            DatasheetDetailsPage datasheetDetailsPage = editFirstService.EditFirstItem();
            editFirstService.Close();
            //datasheetDetailsPage.Go_To_New_Navigate();
            RecipesVariantPage recipesPage = datasheetDetailsPage.EditRecipe();
            recipesPage.Go_To_New_Navigate();
            RecipeGeneralInformationPage recipeGeneralInformationPage = recipesPage.ClickOnGeneralInformationFromRecipeDatasheet();
            Assert.AreEqual(recipeType1, recipeGeneralInformationPage.GetRecipeType(), "le filtre n'est pas pris en compte :Le RecipeType ne sont pas le même");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_RecipeReportV2Print_Result()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool versionPrint = true;
            bool withDatasheet = false;
            string DocFileNamePdfBegin = "Raw Materials by Recipe Report v2";
            string DocFileNameZipBegin = "All_files_";
            string datasheetName1 = TestContext.Properties["Prodman_DatasheetName1"].ToString();
            string datasheetNameMAD = TestContext.Properties["Prodman_DatasheetNameMAD"].ToString();
            string datasheetName2 = TestContext.Properties["Prodman_DatasheetName2"].ToString();
            string datasheetName3 = TestContext.Properties["Prodman_DatasheetName3"].ToString();
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            string firstServiceName = qtyAjustementPage.GetFirstServiceName();
            ServicePricePage servicePricePage = qtyAjustementPage.EditFirstService();
            servicePricePage.Go_To_New_Navigate();
            servicePricePage.UnfoldAll();
            DatasheetDetailsPage datasheetDetailsPage = servicePricePage.EditFirstItem();
            datasheetDetailsPage.Go_To_New_Navigate(2);
            RecipesVariantPage recipesVariantPage = datasheetDetailsPage.EditRecipe();
            RecipeGeneralInformationPage recipeGeneralInformationPage = recipesVariantPage.ClickOnGeneralInformationFromRecipeDatasheet();
            string commercialName1 = recipeGeneralInformationPage.GetCommercialName1();
            string commercialName2 = recipeGeneralInformationPage.GetCommercialName2();
            recipeGeneralInformationPage.Close();
            servicePricePage.Close();
            datasheetDetailsPage.Go_To_New_Navigate(0);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, firstServiceName);
            ResultPage resultsTab = qtyAjustementPage.GoToResultPage();
            DeleteAllFileDownload();
            homePage.ClearDownloads();
            PrintReportPage recipeV2Report = resultsTab.PrintReport(ResultPage.PrintType.RecipeReportV2, versionPrint, withDatasheet);
            recipeV2Report.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            recipeV2Report.Close();
            string trouve = recipeV2Report.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fichier = new FileInfo(trouve);
            fichier.Refresh();
            Assert.IsTrue(fichier.Exists, "Pas de print pdf");
            PdfDocument document = PdfDocument.Open(fichier.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
                foreach (Word mot in p.GetWords())
                    mots.Add(mot.Text);
            bool datasheetExistDansPDF = mots.Contains(datasheetName1) || mots.Contains(datasheetName2) || mots.Contains(datasheetName3) || mots.Contains(datasheetNameMAD);
            Assert.IsFalse(datasheetExistDansPDF, "les datasheet ne doivent pas etre dispo dans le rapport");
            withDatasheet = true;
            resultsTab = qtyAjustementPage.GoToResultPage();
            DeleteAllFileDownload();
            homePage.ClearDownloads();
            PrintReportPage recipeV2ReportWithDatasheet = resultsTab.PrintReport(ResultPage.PrintType.RecipeReportV2, versionPrint, withDatasheet);
            recipeV2ReportWithDatasheet.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            recipeV2ReportWithDatasheet.Close();
            string trouve2 = recipeV2Report.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fichier2 = new FileInfo(trouve2);
            fichier2.Refresh();
            Assert.IsTrue(fichier2.Exists, "Pas de print pdf");
            PdfDocument document2 = PdfDocument.Open(fichier2.FullName);
            mots = new List<string>();
            foreach (Page p in document2.GetPages())
                foreach (var mot in p.GetWords())
                    mots.Add(mot.Text);
            datasheetExistDansPDF = mots.Contains(datasheetName1) || mots.Contains(datasheetName2) || mots.Contains(datasheetName3) || mots.Contains(datasheetNameMAD);
            Assert.IsTrue(datasheetExistDansPDF, "les datasheet ne doivent pas etre dispo dans le rapport");
            bool productionNameExist = (mots.Contains(commercialName1) || mots.Contains(commercialName2));
            Assert.IsTrue(productionNameExist, $"Les production names ({commercialName1} et {commercialName2}) doivent etre pris en compte");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_DatasheetByguest_Result()
        {
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string directory = TestContext.Properties["DownloadsPath"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Datasheet_-_";
            string DocFileNameZipBegin = "All_files_";
            string ProductionServiceName = "TestProductionName";

            // Arrange 
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
          
            IEnumerable<string> allServiceName = qtyAjustementPage.GetAllServiceName();
            string siteSelected = qtyAjustementPage.GetSite();
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.SearchData();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.UnfoldAll();
            Dictionary<string, List<string>> groupAndItems = resultPage.GetGroupAndItemsAssociated();
            List<string> itemsNames = resultPage.GetItemNames();
            var useCaseDetailsList = new List<IEnumerable<KeyValuePair<string, string>>>();
            List<string> selectedGuest = resultPage.GetSelectedGuestType();
            int i = 0;
            foreach (string item in itemsNames)
            {
                ResultModal useCaseModal = resultPage.GoToUseCaseModalPage(item, i);
                useCaseModal.UnfoldAll();
                string serviceNameUseCase = useCaseModal.GetServiceName_UseCase();
                string flightUseCase = useCaseModal.GetFlight_UseCase();
                string dateUseCase = useCaseModal.GetDate_UseCase();
                string customerUseCase = useCaseModal.GetCustomer_UseCase();
                string quatityService = useCaseModal.GetSERVICE_QTE_UseCase();
                string recipeUseCase = useCaseModal.GetRecipe_UseCase();
                // Navigate to the flight page and get the guest type
                FlightPage flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.SetDateState(DateUtils.Now.AddDays(+1));
                string dateFlightString = flightUseCase.Split(' ')[2];
                DateTime dateFlight = DateTime.ParseExact(dateFlightString, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                flightPage.SetDateState(dateFlight);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightUseCase);
                flightPage.Filter(FlightPage.FilterType.Sites, siteSelected);
                FlightDetailsPage editFlight = flightPage.EditFirstFlight(flightUseCase);
                string guestTypeFromFlight = editFlight.GetGuestType();
                editFlight.CloseModal();
                prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
                qtyAjustementPage = prodManagement.DoneToQtyAjustement();
                resultPage = qtyAjustementPage.GoToResultPage();
                resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
                resultPage.UnfoldAll();
                List<KeyValuePair<string, string>> useCaseDetails = new List<KeyValuePair<string, string>>
                    {
                        new KeyValuePair<string, string>("Service", serviceNameUseCase),
                        new KeyValuePair<string, string>("Flight", flightUseCase),
                        new KeyValuePair<string, string>("DateProduction", dateUseCase),
                        new KeyValuePair<string, string>("Customer", customerUseCase),
                        new KeyValuePair<string, string>("QuantityService", quatityService),
                        new KeyValuePair<string, string>("Recipe", recipeUseCase),
                        new KeyValuePair<string, string>("GuestType", guestTypeFromFlight)
                    };
                useCaseDetailsList.Add(useCaseDetails);
                i++;
            }
            prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            qtyAjustementPage.ClearDownloads();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteSelected);
            resultPage = qtyAjustementPage.GoToResultPage();
            List<string> customerselected = resultPage.GetSelectedCustomers();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            List<string> guestSelected = resultPage.GetSelectedGuestType();
            PrintReportPage reportPage = resultPage.PrintReport(ResultPage.PrintType.DatasheetByGuest, true);
            if (reportPage.IsReportGenerated())
            {
                reportPage.Close();
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();

                PdfDocument document = PdfDocument.Open(trouve);
                var pages = document.GetPages().ToList<Page>();
                var words = pages.SelectMany(p => p.GetWords().Select(w => w.Text).ToList()).ToList();

                foreach (var details in useCaseDetailsList)
                {
                    string customerUseCase = details.FirstOrDefault(kvp => kvp.Key == "Customer").Value.ToUpper();
                    string recipeUseCase = details.FirstOrDefault(kvp => kvp.Key == "Recipe").Value;
                    string guestTypeFromFlight = details.FirstOrDefault(kvp => kvp.Key == "GuestType").Value;
                    string dateProductionCase = details.FirstOrDefault(kvp => kvp.Key == "DateProduction").Value;
                    string QtyCase = details.FirstOrDefault(kvp => kvp.Key == "QuantityService").Value;
                    int customerIndex = words.IndexOf(customerUseCase);
                    int serviceIndex = words.IndexOf(ProductionServiceName);
                    int recipeIndex = words.IndexOf(recipeUseCase);
                    int guestIndex = words.IndexOf(guestTypeFromFlight);
                    Assert.AreNotEqual(-1, customerIndex);
                    Assert.AreNotEqual(-1, serviceIndex);
                    Assert.AreNotEqual(-1, recipeIndex);
                    Assert.AreNotEqual(-1, guestIndex);
                    Assert.IsTrue(customerIndex < serviceIndex && serviceIndex < guestIndex && recipeIndex > guestIndex, "le document doit être trié dans l'ordre suivant customer, service name, guest name, par recipe");
                    Assert.IsNotNull(words.IndexOf(dateProductionCase), "Veuillez vérifier la date de production");
                }
                Assert.IsTrue(words.Contains(siteSelected), "The Sites filter was not applied correctly");
            }
            else
            {
                Assert.Fail("Failed Report download");
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_DatasheetByRecipe_Result()
        {
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string directory = TestContext.Properties["DownloadsPath"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Datasheet_-_";
            string DocFileNameZipBegin = "All_files_";
            string ProductionServiceName = "TestProductionName"; 
            //Arrange
            HomePage homePage = LogInAsAdmin();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            homePage.ClearDownloads();
            homePage.PurgeDownloads();
            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            string siteSelected = qtyAjustementPage.GetSite();
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.UnfoldAll();
            Dictionary<string, List<string>> groupAndItems = resultPage.GetGroupAndItemsAssociated();
            List<String> itemsNames = resultPage.GetItemNames();
            var useCaseDetailsList = new List<IEnumerable<KeyValuePair<string, string>>>();
            List<string> selectedGuest = resultPage.GetSelectedGuestType();
            foreach (var item in itemsNames)
            {
                int compteur = 0;
                ResultModal useCaseModal = resultPage.GoToUseCaseModalPage(item, compteur);
                useCaseModal.UnfoldAll();
                string serviceNameUseCase = useCaseModal.GetServiceName_UseCase();
                string flightUseCase = useCaseModal.GetFlight_UseCase();
                string dateUseCase = useCaseModal.GetDate_UseCase();
                string customerUseCase = useCaseModal.GetCustomer_UseCase();
                string quatityService = useCaseModal.GetSERVICE_QTE_UseCase();
                string recipeUseCase = useCaseModal.GetRecipe_UseCase();

                FlightPage flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightUseCase);
                flightPage.Filter(FlightPage.FilterType.Sites, siteSelected);
                FlightDetailsPage editFlight = flightPage.EditFirstFlight(flightUseCase);
                string guestTypeFromFlight = editFlight.GetGuestType();
                editFlight.CloseModal();
                // Prod Man 
                prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
                qtyAjustementPage = prodManagement.DoneToQtyAjustement();
                resultPage = qtyAjustementPage.GoToResultPage();
                resultPage.UnfoldAll();
                var useCaseDetails = new List<KeyValuePair<string, string>>
                    {
                        new KeyValuePair<string, string>("Service", serviceNameUseCase),
                        new KeyValuePair<string, string>("DateProduction", dateUseCase),
                        new KeyValuePair<string, string>("Customer", customerUseCase),
                        new KeyValuePair<string, string>("QuantityService", quatityService),
                        new KeyValuePair<string, string>("Recipe", recipeUseCase),
                        new KeyValuePair<string, string>("GuestType", guestTypeFromFlight)
                    };
                useCaseDetailsList.Add(useCaseDetails);

                compteur++;
            }
            prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            qtyAjustementPage.ClearDownloads();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteSelected);
            resultPage = qtyAjustementPage.GoToResultPage();
            List<string> customerselected = resultPage.GetSelectedCustomers();
            List<string> guestSelected = resultPage.GetSelectedGuestType();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            PrintReportPage reportPage = resultPage.PrintReport(ResultPage.PrintType.DatasheetByRecipe, true);
            if (reportPage.IsReportGenerated())
            {
                reportPage.Close();
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                PdfDocument document = PdfDocument.Open(fi.FullName);
                var pages = document.GetPages().ToList<Page>();
                var words = pages.SelectMany(p => p.GetWords().Select(w => w.Text).ToList()).ToList();
                foreach (var details in useCaseDetailsList)
                {
                    string customerUseCase = details.FirstOrDefault(kvp => kvp.Key == "Customer").Value.ToUpper();
                    string recipeUseCase = details.FirstOrDefault(kvp => kvp.Key == "Recipe").Value;
                    string guestTypeFromFlight = details.FirstOrDefault(kvp => kvp.Key == "GuestType").Value;
                    string dateProductionCase = details.FirstOrDefault(kvp => kvp.Key == "DateProduction").Value;
                    string QtyCase = details.FirstOrDefault(kvp => kvp.Key == "QuantityService").Value;
                    int customerIndex = words.IndexOf(customerUseCase);
                    int serviceIndex = words.IndexOf(ProductionServiceName);
                    int recipeIndex = words.IndexOf(recipeUseCase);
                    int guestIndex = words.IndexOf(guestTypeFromFlight);
                    Assert.AreNotEqual(-1, customerIndex);
                    Assert.AreNotEqual(-1, serviceIndex);
                    Assert.AreNotEqual(-1, recipeIndex);
                    Assert.AreNotEqual(-1, guestIndex);
                    Assert.IsTrue(customerIndex < serviceIndex && serviceIndex < recipeIndex && recipeIndex < guestIndex, "le document doit être trié dans l'ordre suivant customer, service name, recipe name, par guest");
                    Assert.IsNotNull(words.IndexOf(dateProductionCase), "Veuillez vérifier la date de production");
                }
          
                Assert.IsTrue(words.Contains(siteSelected), "le filtre 'Sites' n'est pas appliqué correctement");
            }
            else
            {
                Assert.Fail("Failed Report download");
            }

        }
        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Calcul_OF_Result()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string serviceNameVacuum = TestContext.Properties["Prodman_ServiceVacuum"].ToString();
            string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();
            string datasheetName1 = TestContext.Properties["Prodman_DatasheetName1"].ToString();
            string ProdmanItem1 = TestContext.Properties["Prodman_Item1"].ToString();
            string ProdmanItem2 = TestContext.Properties["Prodman_Item2"].ToString();
            string ProdmanItem3 = TestContext.Properties["Prodman_Item3"].ToString();
            string ProdmanItem4 = TestContext.Properties["Prodman_Item4"].ToString();
            string guestType1 = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();
            string guestType2 = TestContext.Properties["Prodman_Needs_GuestType2"].ToString();
            string customerName1 = TestContext.Properties["Prodman_Customer1"].ToString();
            string customerName2 = TestContext.Properties["Prodman_Customer2"].ToString();
            string customerName3 = TestContext.Properties["Prodman_Customer3"].ToString();
            string placeFrom = "Economato";
            string placeTo = "Produccion";
            string flightNameNormal = "" + TestContext.Properties["Prodman_FlightNormal"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameVacuum = TestContext.Properties["Prodman_FlightVacuum"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string flightNameValide = TestContext.Properties["Prodman_FlightValide"].ToString() + DateUtils.Now.ToString("dd/MM/yyyy");
            string decimalSeparator = ",";

            //Arrange
            HomePage homePage = LogInAsAdmin();
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameNormal);
            FlightDetailsPage editPage = flightPage.EditFirstFlight(flightNameNormal);
            bool isCancelled = editPage.FlightIsCancelled();
            editPage.CloseViewDetails();
            if (isCancelled)
            {
                flightPage.UnSetNewState("C");
                WebDriver.Navigate().Refresh();
                flightPage.SetNewState("P");
            }

            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            prodManagement.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
            prodManagement.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, DateUtils.Now);
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            bool isServiceVacuum = qtyAjustementPage.IsServicePresent(serviceNameVacuum);
            bool isServiceValide = qtyAjustementPage.IsServicePresent(serviceNameFlightValide);
            if (!isServiceNormal || !isServiceVacuum || !isServiceValide)
            {
                prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
                prodManagement.ResetFilter();
                prodManagement.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
                prodManagement.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");
                qtyAjustementPage = prodManagement.DoneToQtyAjustement();
                qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            }
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameNormal), "Le service " + serviceNameNormal + " n'est pas visible.");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameVacuum), "Le service " + serviceNameVacuum + " n'est pas visible.");
            Assert.IsTrue(qtyAjustementPage.IsServicePresent(serviceNameFlightValide), "Le service " + serviceNameFlightValide + " n'est pas visible.");

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowNormalAndVacuumProd, true);

            ResultPage resultsTab = qtyAjustementPage.GoToResultPage();
            resultsTab.UnfoldAll();
            Dictionary<string, List<string>> mapGroupeItems = resultsTab.GetGroupAndItemsAssociated();
            List<string> itemsNames = new List<string>();
            foreach (KeyValuePair<string, List<string>> entry in mapGroupeItems)
            {
                string key = entry.Key;
                List<string> values = entry.Value;
                foreach (string value in values)
                    itemsNames.Add(value);
            }
            qtyAjustementPage = resultsTab.GoToQtyAdjustementPage(); qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameNormal);
            ServicePricePage servicePricePage = qtyAjustementPage.EditFirstService();
            servicePricePage.Go_To_New_Navigate();
            servicePricePage.UnfoldAll();
            servicePricePage.SetFirstPriceDataSheet(datasheetName1);
            DatasheetDetailsPage datasheetDetailsPage = servicePricePage.EditFirstItem();
            servicePricePage.Close();
            datasheetDetailsPage.Go_To_New_Navigate(1);
            RecipesVariantPage recipesVariantPage = datasheetDetailsPage.EditRecipe();
            double grossQty1 = recipesVariantPage.GetGrossQuantity(decimalSeparator);
            ItemGeneralInformationPage itemGeneralInformationPage = recipesVariantPage.EditIngredient(ProdmanItem1);
            itemGeneralInformationPage.Close();
            itemGeneralInformationPage = recipesVariantPage.EditIngredient(ProdmanItem3);
            itemGeneralInformationPage.Close();
            recipesVariantPage.Close();
            recipesVariantPage.WaitPageLoading();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameFlightValide);
            servicePricePage = qtyAjustementPage.EditFirstService();
            servicePricePage.Go_To_New_Navigate();
            servicePricePage.UnfoldAll();
            datasheetDetailsPage = servicePricePage.EditFirstItem();
            servicePricePage.Close();
            datasheetDetailsPage.Go_To_New_Navigate();
            recipesVariantPage = datasheetDetailsPage.EditRecipe();
            double grossQty4 = recipesVariantPage.GetGrossQuantity(decimalSeparator);
            itemGeneralInformationPage = recipesVariantPage.EditIngredient(ProdmanItem4);
            itemGeneralInformationPage.Close();
            recipesVariantPage.Close();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameVacuum);
            servicePricePage = qtyAjustementPage.EditFirstService();
            servicePricePage.Go_To_New_Navigate();
            servicePricePage.UnfoldAll();
            datasheetDetailsPage = servicePricePage.EditFirstItem();
            servicePricePage.Close();
            datasheetDetailsPage.Go_To_New_Navigate();
            recipesVariantPage = datasheetDetailsPage.EditRecipe();
            double grossQty2 = recipesVariantPage.GetGrossQuantity(decimalSeparator);
            itemGeneralInformationPage = recipesVariantPage.EditIngredient(ProdmanItem2);
            itemGeneralInformationPage.Close();
            recipesVariantPage.Close();
            prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, "ALL");
            double qty1 = qtyAjustementPage.GetQuantity(decimalSeparator, serviceNameNormal);
            double qty2 = qtyAjustementPage.GetQuantity(decimalSeparator, serviceNameVacuum);
            double qty3 = qtyAjustementPage.GetQuantity(decimalSeparator, serviceNameFlightValide);
            resultsTab = qtyAjustementPage.GoToResultPage();
            string outputFormNumber = resultsTab.CreateOutputForm(placeFrom, placeTo);
            PageObjects.Warehouse.OutputForm.OutputFormItem outputFormItemPage = resultsTab.GenerateOutputForm();
            double physQuatityOF = outputFormItemPage.GetPhysicalQuantitySum();
            List<string> outputFormItemsNames = outputFormItemPage.GetItemNames();
            Assert.AreEqual(outputFormItemsNames.Count, itemsNames.Count, "the item names are not equal cas 1");
            Assert.IsTrue(!outputFormItemsNames.Except(itemsNames).Any(), "the item names are not equal cas 2");
            double totQty1 = grossQty1 * qty1;
            double totQty2 = grossQty2 * qty2;
            double totQty3 = grossQty4 * qty3;
            double physicalQty = totQty1 + totQty2 + totQty3;
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            double physQuantityOF = physQuatityOF * 10;
            double physicalQuantity = Double.Parse(physicalQty.ToString(), ci);
            double physQuantityOFF = Double.Parse(physQuantityOF.ToString(), ci);
            Assert.AreEqual(physicalQuantity, physQuantityOFF, "Les deux quantités ne sont pas égales");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_RecipeReportdetailedV2_Result()
        {
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string directory = TestContext.Properties["DownloadsPath"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            DateTime dateFrom = DateUtils.Now;
            DateTime dateTo = DateUtils.Now;
            string DocFileNamePdfBegin = "Raw Materials by Recipe Report v2";
            string DocFileNameZipBegin = "All_files_";
            bool versionPrint = true;
            bool withDatasheet = true;
            //Arrange 
            HomePage homePage = LogInAsAdmin();
            var prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateFrom, dateFrom);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, dateFrom);
            string siteSelected = qtyAjustementPage.GetSite();
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.UnfoldAll();
            
            Dictionary<string, List<Dictionary<string, string>>>  itemsWithRow = resultPage.GetGroupWithAllRowMultipleItem();
            List<string> itemsNames = resultPage.GetItemNames();
            List<IEnumerable<KeyValuePair<string, string>>> useCaseDetailsList = new List<IEnumerable<KeyValuePair<string, string>>>();
            foreach (string item in itemsNames)
            {
                int compteur = 0;
                ResultModal useCaseModal = resultPage.GoToUseCaseModalPage(item, compteur);
                useCaseModal.UnfoldAll();
                string serviceNameUseCase = useCaseModal.GetServiceName_UseCase();
                string flightUseCase = useCaseModal.GetFlight_UseCase();
                string dateUseCase = useCaseModal.GetDate_UseCase();
                string customerUseCase = useCaseModal.GetCustomer_UseCase();
                string quantityService = useCaseModal.GetSERVICE_QTE_UseCase();
                string recipeUseCase = useCaseModal.GetRecipe_UseCase();
                string itemQtyUseCase = useCaseModal.GetQTEItem_UseCase();
                string quantityUseCase = useCaseModal.GetQuantity_UseCase();

                // Search for the "workshop" associated with the current item
                string workshop = itemsWithRow
                    .SelectMany(group => group.Value) // Flatten each group's list of dictionaries into a single list
                    .Where(row => row.ContainsKey("item") && row["item"] == item)
                    .Select(row => row.ContainsKey("workshop") ? row["workshop"] : null)
                    .FirstOrDefault();

                useCaseModal.CloseModal();

                List<KeyValuePair<string, string>> useCaseDetails = new List<KeyValuePair<string, string>>
    {
        new KeyValuePair<string, string>("Item", item),
        new KeyValuePair<string, string>("Service", serviceNameUseCase),
        new KeyValuePair<string, string>("Flight", flightUseCase),
        new KeyValuePair<string, string>("DateProduction", dateUseCase),
        new KeyValuePair<string, string>("Customer", customerUseCase),
        new KeyValuePair<string, string>("QuantityService", quantityService),
        new KeyValuePair<string, string>("Recipe", recipeUseCase),
        new KeyValuePair<string, string>("ItemQty", itemQtyUseCase),
        new KeyValuePair<string, string>("Quantity", quantityUseCase),
        new KeyValuePair<string, string>("Workshop", workshop),
    };

                useCaseDetailsList.Add(useCaseDetails);
                compteur++;
            }
            resultPage.ClearDownloads();
            PrintReportPage reportPage = resultPage.PrintReport(ResultPage.PrintType.RecipeReportV2, versionPrint, withDatasheet);
            if (reportPage.IsReportGenerated())
            {
                reportPage.Close();
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                PdfDocument document = PdfDocument.Open(fi.FullName);
                List<string> allWords = new List<string>();
                foreach (Page page in document.GetPages())
                    allWords.AddRange(page.GetWords().Select(w => w.Text));
                foreach (IEnumerable<KeyValuePair<string, string>> details in useCaseDetailsList)
                {
                    string dateProductionCase = details.FirstOrDefault(kvp => kvp.Key == "DateProduction").Value;
                    string customerUseCase = details.FirstOrDefault(kvp => kvp.Key == "Customer").Value;
                    string serviceNameUseCase = details.FirstOrDefault(kvp => kvp.Key == "Service").Value;
                    string recipeUseCase = details.FirstOrDefault(kvp => kvp.Key == "Recipe").Value;
                    string itemUseCase = details.FirstOrDefault(kvp => kvp.Key == "Item").Value;
                    string workshopUseCase = details.FirstOrDefault(kvp => kvp.Key == "Workshop").Value;
                    int customerIndex = reportPage.FindWordIndex(allWords, customerUseCase.ToUpper());
                    int dateProductionIndex = reportPage.FindWordIndex(allWords, dateProductionCase.ToUpper(), customerIndex - 12);
                    int serviceNameIndex = reportPage.FindWordIndex(allWords, serviceNameUseCase.Trim());
                    int recipeIndex = reportPage.FindWordIndex(allWords, recipeUseCase.Trim());
                    int itemIndex = reportPage.FindWordIndex(allWords, itemUseCase.Trim());
                    int workshopIndex = reportPage.FindWordIndex(allWords, workshopUseCase.Trim());
                    Assert.IsTrue(allWords.Count(w => w == dateProductionCase) >= 1, "Veuillez vérifier la date de production");
                    Assert.IsTrue(customerIndex > 0, "Veuillez vérifier le nom du customer"); 
                    Assert.IsTrue(serviceNameIndex > 0, "Veuillez vérifier le nom du service");
                    Assert.IsTrue(recipeIndex > 0, "Veuillez vérifier la procedure de la recette et/ou sous recette");
                    Assert.IsTrue(itemIndex > 0, "Veuillez vérifier le nom de l’item ");
                    Assert.IsTrue(workshopIndex > 0, "Veuillez vérifier le workshop de l’item ");
                }

            }
        }

         [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_RawMatByWorkshop_Result()
        {

            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string directory = TestContext.Properties["DownloadsPath"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            DateTime dateFrom = DateUtils.Now;
            DateTime dateTo = DateUtils.Now;
            string DocFileNamePdfBegin = "Raw materials Report";
            string DocFileNameZipBegin = "All_files_";
            // Arrange 
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateFrom, dateFrom);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, dateFrom);
            string siteSelected = qtyAjustementPage.GetSite();
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowHiddenArticles, true);
            resultPage.UnfoldAll();
            Dictionary<string, List<Dictionary<string, string>>> itemsWithRow = resultPage.GetGroupWithAllRowMultipleItem();
            List<string> customerselected = resultPage.GetSelectedCustomers();
            List<string> guestSelected = resultPage.GetSelectedGuestType();
            resultPage.ClearDownloads();
            PrintReportPage reportPage = resultPage.PrintReport(ResultPage.PrintType.RawMaterialsByWorkshop, true);
            if (reportPage.IsReportGenerated())
            {
                reportPage.Close();
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                PdfDocument document = PdfDocument.Open(fi.FullName);
                string allWords = "";
                foreach (Page page in document.GetPages())
                    allWords += page.Text.Replace(" ", "");
                foreach (KeyValuePair<string, List<Dictionary<string, string>>> details in itemsWithRow)
                {
                    string workshop = details.Value.FirstOrDefault(dic => dic.ContainsKey("workshop"))?["workshop"];
                    string item = details.Value.FirstOrDefault(dic => dic.ContainsKey("item"))?["item"];
                    string totProdQty = details.Value.FirstOrDefault(dic => dic.ContainsKey("quantity"))?["quantity"];
                    string packaging = details.Value.FirstOrDefault(dic => dic.ContainsKey("packaging"))?["packaging"];

                    bool workshopIndex = allWords.Contains(workshop.Replace(" ", ""));
                    bool itemIndex = allWords.Contains(item.Replace(" ", ""));
                    bool packagingIndex = allWords.Contains(packaging == "bolsa 1Kg" ? "KG" : packaging.Replace(" ", ""));
                    homePage.Navigate();
                    ItemPage purchasing_ItemPage = homePage.GoToPurchasing_ItemPage();
                    purchasing_ItemPage.ResetFilter();
                    purchasing_ItemPage.Filter(ItemPage.FilterType.Search, item);
                    purchasing_ItemPage.UnfoldAll();
                    ItemGeneralInformationPage itemGeneralInformationPage = purchasing_ItemPage.ClickOnFirstItem();
                    decimal storageQty = itemGeneralInformationPage.GetFirstPackagingStorageQty("");
                    string storageUnit = itemGeneralInformationPage.GetFirstPackagingStorUnit();
                    bool totalStorageQtyIndex = allWords.Contains(packaging.Replace(" ", ""));
                    string totProdQtyExtracted = new string(totProdQty.Where(c => char.IsDigit(c) || c == ',' || c == '.').ToArray());
                    decimal TotStorageQty = decimal.Parse(totProdQtyExtracted, System.Globalization.CultureInfo.GetCultureInfo("fr-FR")) / storageQty;
                    string numbersAfterDecimal = TotStorageQty.ToString(System.Globalization.CultureInfo.InvariantCulture).Contains('.')
                    ? TotStorageQty.ToString(System.Globalization.CultureInfo.InvariantCulture).Split('.')[1]
                    : string.Empty;
                    bool qtyIndex = allWords.Contains(numbersAfterDecimal);
                    bool storageUnitIndex = allWords.Contains(storageUnit.Replace(" ", ""));
                    Assert.IsTrue(workshopIndex && itemIndex, " les items ne sont pas réparti par workshop");
                    Assert.IsTrue(itemIndex, "le nom de l'item du report ne correspond pas au nom de l'item dans l'index Result");
             //       Assert.IsTrue(qtyIndex, " la somme du Total prod qty du rapport ne correspond pas au somme des Qty dans l'index Result");
                    Assert.IsTrue(storageUnitIndex, "Le storage unit du rapport ne correspond pas au storage unit dans le module purchasing/item");
                    Assert.IsTrue(packagingIndex, "le Prod unit du rapport ne correspond pas au packaging dans l'index Result");
                }
                foreach (string customer in customerselected)
                {
                    if (!allWords.Contains(customer.Replace(" ", "")))
                        Assert.Fail("echec customer " + customer);
                }
                int indexOfFirst = allWords.IndexOf("GuestTypes");
                if (indexOfFirst != -1)
                {
                    string substringAfterFirst = allWords.Substring(indexOfFirst + "GuestTypes".Length);
                    if (!substringAfterFirst.Contains("ALL"))
                        Assert.Fail("not ALL guest selected");
                }
                bool isContainingSelectedSite = allWords.Contains(siteSelected);
                Assert.IsTrue(isContainingSelectedSite, "le filtre 'Sites' n'est pas appliqué correctement");

            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrevalidateOrValidateFlightOnly()
        {
            // Pour ce test, nous avons besoin d'un seul service lié à un vol , on debut le vol sera pre val puis val .
            // L'objectif de ce test est de verifier que les services des vols pre val et val serons affichers sur QA meme apres changement de status . 

            // Prepare
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            string flightNameNormal = "" + TestContext.Properties["Prodman_FlightNormal"].ToString() + DateTime.Today.AddDays(1).ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();
            try
            {
                FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
                prodManagement.ResetFilter();
                QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
                
                // Assert 1 : Verif que le service liee au vol pre val est afficher 
                bool IsServicePresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);
                Assert.IsTrue(IsServicePresent, $"Le service {serviceNameNormal} n'existe pas , le vol pre val n est pas visible");
                
                FlightPage flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameNormal); 
                flightPage.SetDateState(DateUtils.Now.AddDays(+1));
                flightPage.SetNewState("V");
                prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
                prodManagement.ResetFilter();
                qtyAjustementPage = prodManagement.DoneToQtyAjustement();
                // Assert 2 : Verif que le service liee au vol VAL est afficher 
                IsServicePresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);
                Assert.IsTrue(IsServicePresent, $"Le service {serviceNameNormal} ne doit pas etre existe , le vol val n est pas visible ");
            }

            finally
            {
                FlightPage flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNameNormal);
                flightPage.SetDateState(DateUtils.Now.AddDays(1));
                string preval = "Preval";
                flightPage.Filter(FlightPage.FilterType.Status, preval);
                flightPage.UnSetNewState("V");
                FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
                prodManagement.ResetFilter();
                QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
                bool IsServicePresent = qtyAjustementPage.IsServicePresent(serviceNameNormal);

                Assert.IsTrue(IsServicePresent, $"Le service {serviceNameNormal} n'existe pas , le vol pre val n est pas visible (2eme test)");
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Filter_ItemHide_Result()
        {
           
            int resultsCountBeforeRefresh = 0;
            int resultsCountAfterRefresh = 0;
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultsCountBeforeRefresh = resultPage.CountResults();
            if (resultsCountBeforeRefresh <= 0)
                 qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowHiddenArticles, true);
            resultPage.UnfoldAll();
           
            try
            {
                resultPage.HideFirstArticle();
                WebDriver.Navigate().Refresh();
                resultPage = qtyAjustementPage.GoToResultPage();
                resultPage = qtyAjustementPage.GoToResultPage();
                bool isHiddenItemsCounterAlertActive = resultPage.IsHiddenItemsCounterAlertActive();
                Assert.IsTrue(isHiddenItemsCounterAlertActive, "Hidden Items Counter Alert n'est pas visible");

                resultsCountAfterRefresh = resultPage.CountResults();
            }
           finally
            {
                resultPage = qtyAjustementPage.GoToResultPage();
                int resultsCount = resultPage.CountResults();
                if (resultsCount <= 0)
                    qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowHiddenArticles, true);
                    resultPage.UnfoldAll();
                    resultPage.HideFirstArticle();
            }

            Assert.AreNotEqual(resultsCountAfterRefresh, resultsCountBeforeRefresh, "Le nombre d’articles n’est pas égal.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_AssemblyReportServiceName_Result()
        {
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();

            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Assembly Report_-_";
            string DocFileNameZipBegin = "All_files_";
            string productionName = "Test Production Name";
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.ClearDownloads();

            // Report check service name 
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameNormal);
            string returnService = qtyAjustementPage.GetServiceName();
            ServicePricePage first_service = qtyAjustementPage.EditFirstService();
            ServiceGeneralInformationPage general_info = first_service.ClickOnGeneralInformationTabOtherNavigation();
            string production_name = general_info.GetProductionName();
            filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameNormal);
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            PrintReportPage reportPage = resultPage.PrintReport(ResultPage.PrintType.AssemblyReport, true, false, false);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            PdfDocument documentservicename = PdfDocument.Open(fi.FullName);
            string allWords = "";
            foreach (Page p in documentservicename.GetPages())
            {
                allWords += p.Text.Replace(" ", "");
            }
            bool containsServiceProdmanNormal = allWords.Contains("ServiceProdmanNormal");
            // Assert 1 : Service Name 
            Assert.IsTrue(containsServiceProdmanNormal, "The service name is not available in the PDF");

            // Report check production name
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameNormal);
            resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.ClearDownloads();
            reportPage = resultPage.PrintReport(ResultPage.PrintType.AssemblyReport, true, false, true);
            //reportPage.Close();
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fiProduction = new FileInfo(trouve);
            fiProduction.Refresh();
            PdfDocument documentProductionname = PdfDocument.Open(fiProduction.FullName);
            allWords = "";
            foreach (Page p in documentProductionname.GetPages())
            {
                allWords += p.Text.Replace(" ", "");
            }
            productionName = productionName.Replace(" ", "");
            bool containsTestProductionName = allWords.Contains("TestProductionName");
            // Assert 2 : Production Name 
            Assert.IsTrue(containsTestProductionName, "The Production Name is not avalable in the PDF ");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_TraySetUp_Result()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "HACCP Tray Setup_";
            string DocFileNameZipBegin = "All_files_";

            // Arrange 
            HomePage homePage = LogInAsAdmin();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            DateTime dateFromFilter = qtyAjustementPage.GetDateFrom();
            DateTime dateToFilter = qtyAjustementPage.GetDateTo();
            var startTime = qtyAjustementPage.GetStartTime();
            var endTime = qtyAjustementPage.GetEndTime();
            var serviceName = qtyAjustementPage.GetFirstServiceName();
            qtyAjustementPage.ClearDownloads();
            var resultPage = qtyAjustementPage.GoToResultPage();

            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPTraySetup, true);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            reportPage.Close();
            var trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fiProduction = new FileInfo(trouve);
            fiProduction.Refresh();
            PdfDocument documentProductionname = PdfDocument.Open(fiProduction.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in documentProductionname.GetPages())
            {
                mots.AddRange(p.GetWords().Select(m => m.Text));
            }

            var x = mots;
            Assert.IsTrue(mots.Contains(serviceName));
            Assert.IsTrue(mots.Contains(endTime));
            Assert.IsTrue(mots.Contains(startTime));
            string dateOfService = dateFromFilter.ToString("dd/MM/yyyy");
            Assert.IsTrue(mots.Contains(dateOfService));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_TraySetUpMultiFlight_Result()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "HACCP Tray Setup [MultiFlight]_-_";
            string DocFileNameZipBegin = "All_files_";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            DateTime dateFromFilter = qtyAjustementPage.GetDateFrom();
            DateTime dateToFilter = qtyAjustementPage.GetDateTo();
            var startTime = qtyAjustementPage.GetStartTime();
            var endTime = qtyAjustementPage.GetEndTime();
            var serviceName = qtyAjustementPage.GetFirstServiceName();
            qtyAjustementPage.ClearDownloads();
            var resultPage = qtyAjustementPage.GoToResultPage();

            var reportPage = resultPage.PrintReport(ResultPage.PrintType.Haccptraysetupmultiflight, true);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            reportPage.Close();
            var trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fiProduction = new FileInfo(trouve);
            fiProduction.Refresh();
            PdfDocument documentProductionname = PdfDocument.Open(fiProduction.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in documentProductionname.GetPages())
            {
                mots.AddRange(p.GetWords().Select(m => m.Text));
            }

            Assert.IsTrue(mots.Contains(serviceName));
            Assert.IsTrue(mots.Contains(endTime));
            Assert.IsTrue(mots.Contains(startTime));
            string dateOfService = dateFromFilter.ToString("dd/MM/yyyy");
            Assert.IsTrue(mots.Contains(dateOfService));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_TraySetUpMultiFlightNoGuest_Result()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "HACCP Tray Setup [MultiFlight no Guest]_-_";
            string DocFileNameZipBegin = "All_files_";

            // Arrange 
            HomePage homePage = LogInAsAdmin();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            DateTime dateFromFilter = qtyAjustementPage.GetDateFrom();
            DateTime dateToFilter = qtyAjustementPage.GetDateTo();
            var startTime = qtyAjustementPage.GetStartTime();
            var endTime = qtyAjustementPage.GetEndTime();
            var serviceName = qtyAjustementPage.GetFirstServiceName();
            qtyAjustementPage.ClearDownloads();
            var resultPage = qtyAjustementPage.GoToResultPage();

            var reportPage = resultPage.PrintReport(ResultPage.PrintType.Haccptraysetupmultiflightnoguest, true);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            reportPage.Close();
            var trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fiProduction = new FileInfo(trouve);
            fiProduction.Refresh();
            PdfDocument documentProductionname = PdfDocument.Open(fiProduction.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in documentProductionname.GetPages())
            {
                mots.AddRange(p.GetWords().Select(m => m.Text));
            }

            Assert.IsTrue(mots.Contains(serviceName));
            Assert.IsTrue(mots.Contains(endTime));
            Assert.IsTrue(mots.Contains(startTime));
            string dateOfService = dateFromFilter.ToString("dd/MM/yyyy");
            Assert.IsTrue(mots.Contains(dateOfService));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_HaccpPlatingGroupByRecipe_Result()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var prodmanCustomer1 = TestContext.Properties["Prodman_Customer1"].ToString();
            var prodmanCustomer2 = TestContext.Properties["Prodman_Customer2"].ToString();
            var prodmanCustomer3 = TestContext.Properties["Prodman_Customer3"].ToString();
            var recetteProdman1 = TestContext.Properties["Prodman_RecipeName1"].ToString();
            var recetteProdman2 = TestContext.Properties["Prodman_RecipeName2"].ToString();
            var recetteProdman3 = TestContext.Properties["Prodman_RecipeName3"].ToString();
            string DocFileNamePdfBegin = "HACCP Plating [GroupByRecipe]_-_";
            string DocFileNameZipBegin = "All_files_";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            var serviceName = qtyAjustementPage.GetFirstServiceName();
            DateTime dateFromFilter = qtyAjustementPage.GetDateFrom();
            qtyAjustementPage.ClearDownloads();
            var resultPage = qtyAjustementPage.GoToResultPage();
            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPPlatingGroupByRecipe, true);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            reportPage.Close();
            var trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            FileInfo fiProduction = new FileInfo(trouve);
            fiProduction.Refresh();
            PdfDocument documentProductionname = PdfDocument.Open(fiProduction.FullName);
            //List<string> mots = new List<string>();
         
            PdfDocument document = PdfDocument.Open(trouve);
            var pages = document.GetPages().ToList<Page>();
            var allWords = pages.SelectMany(p => p.GetWords().Select(w => w.Text).ToList()).ToList();

            string dateOfService = dateFromFilter.ToString("dd/MM/yyyy");
            Assert.IsTrue(allWords.Contains(recetteProdman1.Replace(" ", "")));
            Assert.IsTrue(allWords.Contains(dateOfService.Replace(" ", "")));
            Assert.IsTrue(allWords.Contains(serviceName.Replace(" ", "")));
            Assert.IsTrue(allWords.Contains(prodmanCustomer1.ToUpper().Replace(" ", "")));
            //int recetteProdman1Index = reportPage.FindWordIndex(mots, recetteProdman1.Trim());
            //int recetteProdman2Index = reportPage.FindWordIndex(mots, recetteProdman2.Trim());
            //int recetteProdman3Index = reportPage.FindWordIndex(mots, recetteProdman3.Trim());
            //string qtyRecetteProdman1 = mots[recetteProdman1Index];
            //string qtyRecetteProdman2 = mots[recetteProdman2Index];
            //string qtyRecetteProdman3 = mots[recetteProdman3Index];
            //int totalQty = int.Parse(qtyRecetteProdman1) + int.Parse(qtyRecetteProdman2) + int.Parse(qtyRecetteProdman3);
           
            //Assert.IsTrue(allWords.Contains(recetteProdman2.Replace(" ", "")));
            //Assert.IsTrue(allWords.Contains(recetteProdman3.Replace(" ", "")));
         
            // Assert.IsTrue(allWords.Contains(prodmanCustomer2.Replace(" ", "")));
            //Assert.IsTrue(allWords.Contains(prodmanCustomer3.Replace(" ", "")));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_HaccpThawing_Result()
        {
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "HACCP Thawing_-_";
            string DocFileNameZipBegin = "All_files_";
            bool versionPrint = true;
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);

            homePage.Navigate();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            var resultPage = qtyAjustementPage.GoToResultPage();
            Assert.IsTrue(resultPage.CountResults() > 0, "Aucune Données de production à générer");
            resultPage.UnfoldAll();
            var itemsWithRow = resultPage.GetGroupWithAllRow();
            var itemsNames = resultPage.GetItemNames();
            var useCaseDetailsList = new List<IEnumerable<KeyValuePair<string, string>>>();
            foreach (var item in itemsNames)
            {
                int compteur = 0;
                ResultModal useCaseModal = resultPage.GoToUseCaseModalPage(item, compteur);
                useCaseModal.UnfoldAll();
                var serviceNameUseCase = useCaseModal.GetServiceName_UseCase();
                var flightUseCase = useCaseModal.GetFlight_UseCase();
                var dateUseCase = useCaseModal.GetDate_UseCase();
                var groupe = itemsWithRow.Where(x => x.Value.ContainsKey("item") && x.Value["item"] == item)
                                    .Select(x => x.Key)
                                    .FirstOrDefault();
                useCaseModal.CloseModal();
                var useCaseDetails = new List<KeyValuePair<string, string>>
        {
            new KeyValuePair<string, string>("Item", item),
            new KeyValuePair<string, string>("Service", serviceNameUseCase),
            new KeyValuePair<string, string>("Flight", flightUseCase),
            new KeyValuePair<string, string>("DateProduction", dateUseCase),
            new KeyValuePair<string, string>("Groupe", groupe),
        };
                useCaseDetailsList.Add(useCaseDetails);

                compteur++;
            }
            homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var itemsPage = homePage.GoToPurchasing_ItemPage();
            foreach (var item in itemsNames)
            {
                itemsPage.ResetFilter();
                itemsPage.Filter(ItemPage.FilterType.Search, item);
                ItemGeneralInformationPage itemGeneralInformation = itemsPage.ClickOnFirstItem();
                if (!itemGeneralInformation.isChecked("IsThawing"))
                {
                    useCaseDetailsList.RemoveAll(details => details.Any(kvp => kvp.Key == "Item" && kvp.Value == item));
                }
                itemGeneralInformation.BackToList();
            }
            homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            resultPage = qtyAjustementPage.GoToResultPage();
            var guestTypeFilterApply = resultPage.GetSelectedGuestType();
            var customersFilterApply = resultPage.GetSelectedCustomers();
            var workshopFilterApply = "ALL"; //resultPage.GetSelectedOptions(ResultPage.ElementIdFilter.SelectedWorkshopsTrueValue);
            var serviceFilterApply = resultPage.GetSelectedOptions(ResultPage.ElementIdFilter.SelectedServicesTrueValue);
            var serviceCategoriesFilterApply = "ALL"; //resultPage.GetSelectedOptions(ResultPage.ElementIdFilter.SelectedServiceCategoriesTrueValue);
            var recipeTypesFilterApply = "ALL"; //resultPage.GetSelectedOptions(ResultPage.ElementIdFilter.SelectedRecipeTypesTrueValue);

            resultPage.ClearDownloads();

            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPThawing, versionPrint);
            if (reportPage.IsReportGenerated())
            {
                reportPage.Close();
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                PdfDocument document = PdfDocument.Open(fi.FullName);
                var allWords = new List<string>();
                var allText = string.Empty;
                foreach (Page page in document.GetPages())
                {
                    allWords.AddRange(page.GetWords().Select(w => w.Text));
                    allText = page.Text;
                }
                foreach (var details in useCaseDetailsList)
                {
                    string dateProductionCase = details.FirstOrDefault(kvp => kvp.Key == "DateProduction").Value;
                    string serviceNameUseCase = details.FirstOrDefault(kvp => kvp.Key == "Service").Value;
                    string flightNameUseCase = details.FirstOrDefault(kvp => kvp.Key == "Flight").Value;
                    string itemUseCase = details.FirstOrDefault(kvp => kvp.Key == "Item").Value;
                    string groupe = details.FirstOrDefault(kvp => kvp.Key == "Groupe").Value;
                    int itemIndex = reportPage.FindWordIndex(allWords, itemUseCase.Trim());
                    int groupeIndex = reportPage.FindWordIndex(allWords, groupe.Trim(), reportPage.FindWordIndex(allWords, "By"));
                    Assert.IsTrue(customersFilterApply.Any(w => allText.Contains(w)), "le filtre 'Customers' n'est pas appliqué correctement");
                    Assert.IsTrue(guestTypeFilterApply.Any(w => allText.Contains(w)), "le filtre 'Guest Types' n'est pas appliqué correctement");
                    Assert.IsTrue( allText.Contains(workshopFilterApply), "le filtre 'Workshops' n'est pas appliqué correctement");
                    Assert.IsTrue(serviceFilterApply.Any(w => allText.Contains(w)), "le filtre 'Services' n'est pas appliqué correctement");
                    Assert.IsTrue( allText.Contains(serviceCategoriesFilterApply), "le filtre 'Service Categories' n'est pas appliqué correctement");
                    Assert.IsTrue(allText.Contains(recipeTypesFilterApply), "le filtre 'Recipe Types' n'est pas appliqué correctement");
                    Assert.IsTrue(allWords.Count(w => w == dateProductionCase) >= 1, "Veuillez vérifier la date de production");
                    //Assert.AreEqual(allWords[itemIndex - 1], groupe, " les items doient être classés par groups d’items");
                    Assert.IsTrue(itemIndex > 0, "Veuillez vérifier le nom de l’item ");

                }
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PrintCheck_HaccpSanitization_Result()
        {
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "HACCP Sanitization_-_";
            string DocFileNameZipBegin = "All_files_";
            bool versionPrint = true;
            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            var resultPage = qtyAjustementPage.GoToResultPage();
            Assert.IsTrue(resultPage.CountResults() > 0, "Aucune Données de production à générer");
            resultPage.UnfoldAll();
            var itemsWithRow = resultPage.GetGroupWithAllRow();
            var itemsNames = resultPage.GetItemNames();
            var useCaseDetailsList = new List<IEnumerable<KeyValuePair<string, string>>>();
            foreach (var item in itemsNames)
            {
                int compteur = 0;
                ResultModal useCaseModal = resultPage.GoToUseCaseModalPage(item, compteur);
                useCaseModal.UnfoldAll();
                var serviceNameUseCase = useCaseModal.GetServiceName_UseCase();
                var flightUseCase = useCaseModal.GetFlight_UseCase();
                var dateUseCase = useCaseModal.GetDate_UseCase();
                var groupe = itemsWithRow.Where(x => x.Value.ContainsKey("item") && x.Value["item"] == item)
                                    .Select(x => x.Key)
                                    .FirstOrDefault();
                var itemQty = useCaseModal.GetQTEItem_UseCase();
                useCaseModal.CloseModal();
                var useCaseDetails = new List<KeyValuePair<string, string>>
        {
            new KeyValuePair<string, string>("Item", item),
            new KeyValuePair<string, string>("ItemQty", itemQty),
            new KeyValuePair<string, string>("DateProduction", dateUseCase),
            new KeyValuePair<string, string>("Groupe", groupe),
        };
                useCaseDetailsList.Add(useCaseDetails);

                compteur++;
            }
            homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var itemsPage = homePage.GoToPurchasing_ItemPage();
            foreach (var item in itemsNames)
            {
                itemsPage.ResetFilter();
                itemsPage.Filter(ItemPage.FilterType.Search, item);
                ItemGeneralInformationPage itemGeneralInformation = itemsPage.ClickOnFirstItem();
                if (!itemGeneralInformation.isChecked("IsSanitization"))
                {
                    useCaseDetailsList.RemoveAll(details => details.Any(kvp => kvp.Key == "Item" && kvp.Value == item));
                }
                itemGeneralInformation.BackToList();
            }
            homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            resultPage = qtyAjustementPage.GoToResultPage();
            string guestTypeFilterApply = "ALL"; // resultPage.GetSelectedGuestType(); 
            var customersFilterApply = resultPage.GetSelectedCustomers();
            string workshopFilterApply = "ALL"/*resultPage.GetSelectedOptions(ResultPage.ElementIdFilter.SelectedWorkshopsTrueValue)*/;
            var serviceFilterApply = resultPage.GetSelectedOptions(ResultPage.ElementIdFilter.SelectedServicesTrueValue);
            string serviceCategoriesFilterApply = "ALL"; //resultPage.GetSelectedOptions(ResultPage.ElementIdFilter.SelectedServiceCategoriesTrueValue);
            string recipeTypesFilterApply = "ALL"; //  resultPage.GetSelectedOptions(ResultPage.ElementIdFilter.SelectedRecipeTypesTrueValue);

            resultPage.ClearDownloads();

            var reportPage = resultPage.PrintReport(ResultPage.PrintType.HACCPSanitization, versionPrint);
            if (reportPage.IsReportGenerated())
            {
                reportPage.Close();
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                PdfDocument document = PdfDocument.Open(fi.FullName);
                var allWords = new List<string>();
                var allText = string.Empty;
              
                var pages = document.GetPages().ToList<Page>();
                var words = pages.SelectMany(p => p.GetWords().Select(w => w.Text).ToList()).ToList();
                foreach (var details in useCaseDetailsList)
                {
                    string dateProductionCase = details.FirstOrDefault(kvp => kvp.Key == "DateProduction").Value;
                    string itemUseCase = details.FirstOrDefault(kvp => kvp.Key == "Item").Value;
                    string itemQtyUseCase = details.FirstOrDefault(kvp => kvp.Key == "ItemQty").Value;
                    //string groupe = details.FirstOrDefault(kvp => kvp.Key == "Groupe").Value;
                    int itemIndex =words.IndexOf(itemUseCase);
                    //int groupeIndex = reportPage.FindWordIndex(words, groupe.Split(' ')[1].Trim());
                    int itemQtyIndex = reportPage.FindWordIndex(words, itemQtyUseCase.Trim(), reportPage.FindWordIndex(words, itemUseCase));
                    Assert.IsTrue( words.Contains(guestTypeFilterApply), "le filtre 'Guest Types' n'est pas appliqué correctement");
                    Assert.IsTrue(words.Contains(workshopFilterApply), "le filtre 'Workshops' n'est pas appliqué correctement");
                    Assert.IsTrue(serviceFilterApply.Any(w => words.Contains(w)), "le filtre 'Services' n'est pas appliqué correctement");
                    Assert.IsTrue( words.Contains(serviceCategoriesFilterApply), "le filtre 'Service Categories' n'est pas appliqué correctement");
                    Assert.IsTrue( words.Contains(recipeTypesFilterApply), "le filtre 'Recipe Types' n'est pas appliqué correctement");
                    Assert.IsTrue(words.Count(w => w == dateProductionCase) >= 1, "Veuillez vérifier la date de production");
                    Assert.IsTrue(itemIndex > 0, "Veuillez vérifier le nom de l’item ");
                    //Assert.IsTrue(itemQtyIndex > 0, "La somme des quantités de l'item n'existe pas. Veuillez vérifier");
                }
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Calcul_SO_Result_Price_KG()
        {
            // Pour ce test, On a pas besoin de service mais on est besoin d un item pour faire le calcul SO   .
            // L'objectif de ce test est de verifier le calcul de Price kg = (pack.price/stor qty)x 1000 / prod qty   . 

            string ChooseOneOption = "Produccion";
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            // Arrange 
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
            // 1/ Generate SO 
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            string supplyOrderNoOptionModal = resultPage.CreateSupplyOrder(null);
            resultPage.ChooseDeliveryLocationOption(ChooseOneOption);
            SupplyOrderItem supplyOrderNoOption = resultPage.GenerateSupplyOrder(true);
            resultPage.Close();
            // Get price/KG from supply orders 
            string price = supplyOrderNoOption.GetPrice_KG();
            string itemName = supplyOrderNoOption.GetFirstItemName();
            // move to the item interface 
            
            ItemGeneralInformationPage itemPageItem = supplyOrderNoOption.EditItemForProdMan(itemName);
            itemPageItem.SearchPackagingProdMan(siteACE);
            string prod_qty_str = itemPageItem.GetFirstPackagingProdQty();
            decimal store_qty = itemPageItem.GetFirstPackagingStorageQty("");
            string unit_price_str = itemPageItem.GetFirstPackagingUnitPriceWithClick();

            // Price/Kg = (pack.price/stor qty)x 1000 / prod qty 
            decimal unit_price = decimal.Parse(unit_price_str);
            decimal prod_qty = decimal.Parse(prod_qty_str); 
            decimal price_from_item = ( (unit_price / store_qty) * 1000 ) / prod_qty;
            // cleaning price from supply order 
            string cleanedPrice = price.Replace("€", "").Trim();
            decimal price_from_SO = decimal.Parse(cleanedPrice);
            Assert.AreEqual(price_from_SO, price_from_item);
        }
        //[Timeout(_timeout)]
        //[TestMethod]
        //public void PR_PRODMAN_Calcul_SO_Result()
        //{
        //    string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
        //    string serviceNameVacuum = TestContext.Properties["Prodman_ServiceVacuum"].ToString();
        //    string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();
        //    string deliveryLocation = "Produccion";
        //    string decimalSeparator = ",";

        //    // Arrange
        //    HomePage homePage = LogInAsAdmin();
        //    FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
        //    prodManagement.ResetFilter();

        //    QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
        //    ResultPage resultsTab = qtyAjustementPage.GoToResultPage();
        //    resultsTab.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
        //    resultsTab.UnfoldAll();
        //    Dictionary<string, List<string>> mapGroupeItems = resultsTab.GetGroupAndItemsAssociated();
        //    List<string> itemsNames = new List<string>();
        //    foreach (KeyValuePair<string, List<string>> entry in mapGroupeItems)
        //    {
        //        string key = entry.Key;
        //        List<string> values = entry.Value;
        //        foreach (string value in values)
        //            itemsNames.Add(value);
        //    }

        //    qtyAjustementPage = resultsTab.GoToQtyAdjustementPage();

        //    /* service name normal */
        //    qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameNormal);
        //    ServicePricePage servicePricePage = qtyAjustementPage.EditFirstService();
        //    servicePricePage.Go_To_New_Navigate();
        //    servicePricePage.UnfoldAll();
        //    servicePricePage.WaitPageLoading();
        //    DatasheetDetailsPage datasheetDetailsPage = servicePricePage.EditFirstItem();
        //    servicePricePage.Close();
        //    datasheetDetailsPage.Go_To_New_Navigate();
        //    RecipesVariantPage recipesVariantPage = datasheetDetailsPage.EditRecipe();
        //    double grossQty1 = recipesVariantPage.GetGrossQuantity(decimalSeparator);
        //    recipesVariantPage.Close();

        //    /* service name Vaccum*/
        //    qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameVacuum);
        //    servicePricePage = qtyAjustementPage.EditFirstService();
        //    servicePricePage.Go_To_New_Navigate();
        //    servicePricePage.UnfoldAll();
        //    datasheetDetailsPage = servicePricePage.EditFirstItem();
        //    servicePricePage.Close();
        //    datasheetDetailsPage.Go_To_New_Navigate();
        //    recipesVariantPage = datasheetDetailsPage.EditRecipe();
        //    double grossQty2 = recipesVariantPage.GetGrossQuantity(decimalSeparator);
        //    recipesVariantPage.Close();

        //    /* service name valid */
        //    qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameFlightValide);
        //    servicePricePage = qtyAjustementPage.EditFirstService();
        //    servicePricePage.Go_To_New_Navigate();
        //    servicePricePage.UnfoldAll();
        //    servicePricePage.WaitPageLoading();
        //    datasheetDetailsPage = servicePricePage.EditFirstItem();
        //    servicePricePage.Close();
        //    datasheetDetailsPage.Go_To_New_Navigate();
        //    recipesVariantPage = datasheetDetailsPage.EditRecipe();
        //    double grossQty3 = recipesVariantPage.GetGrossQuantity(decimalSeparator);
        //    recipesVariantPage.Close();

        //    prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
        //    prodManagement.ResetFilter();
        //    qtyAjustementPage = prodManagement.DoneToQtyAjustement();
        //    qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, "ALL");
        //    double qty1 = qtyAjustementPage.GetQuantity(decimalSeparator, serviceNameNormal);
        //    double qty2 = qtyAjustementPage.GetQuantity(decimalSeparator, serviceNameVacuum);
        //    double qty3 = qtyAjustementPage.GetQuantity(decimalSeparator, serviceNameFlightValide);
        //    resultsTab = qtyAjustementPage.GoToResultPage();
        //    string supplyOrderNumber = resultsTab.CreateSupplyOrder(deliveryLocation);
        //    SupplyOrderItem supplyOrderItemPage = resultsTab.GenerateSupplyOrder();
        //    double physQuatityOF = supplyOrderItemPage.GetProdQuantitySum();
        //    List<string> outputFormItemsNames = supplyOrderItemPage.GetItemNames();
        //    Assert.AreEqual(outputFormItemsNames.Count, itemsNames.Count, "the item names are not equal cas 1");
        //    Assert.IsTrue(!outputFormItemsNames.Except(itemsNames).Any(), "the item names are not equal cas 2");
        //    double totQty1 = grossQty1 * qty1;
        //    double totQty2 = grossQty2 * qty2;
        //    double totQty3 = grossQty3 * qty3;
        //    double physicalQty = totQty1 + totQty2 + totQty3;

        //    CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
        //    double physQuantityOF = physQuatityOF * 10;
        //    double physicalQuantity = Double.Parse(physicalQty.ToString(), ci);
        //    double physQuantityOFF = Double.Parse(physQuantityOF.ToString(), ci);
        //    Assert.AreEqual(physicalQuantity, physQuantityOFF, "Les deux quantités ne sont pas égales");
        //}

        //[TestMethod]
        //public void PR_PRODMAN_Calcul_SO_Result()
        //{
        //    string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
        //    string serviceNameVacuum = TestContext.Properties["Prodman_ServiceVacuum"].ToString();
        //    string serviceNameFlightValide = TestContext.Properties["Prodman_ServiceFlightValide"].ToString();
        //    string deliveryLocation = "Produccion";
        //    LogInAsAdmin();
        //    HomePage homePage = new HomePage(WebDriver, TestContext);
        //    homePage.Navigate();
        //    FilterAndFavoritesPage prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
        //    prodManagement.ResetFilter();
        //    QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();
        //    ResultPage resultsTab = qtyAjustementPage.GoToResultPage();
        //    resultsTab.UnfoldAll();
        //    Dictionary<string, List<string>> mapGroupeItems = resultsTab.GetGroupAndItemsAssociated();
        //    List<string> itemsNames = new List<string>();
        //    foreach (KeyValuePair<string, List<string>> entry in mapGroupeItems)
        //    {
        //        string key = entry.Key;
        //        List<string> values = entry.Value;
        //        foreach (string value in values) itemsNames.Add(value);
        //    }
        //    qtyAjustementPage = resultsTab.GoToQtyAdjustementPage();
        //    qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameNormal);
        //    ServicePricePage servicePricePage = qtyAjustementPage.EditFirstService();
        //    servicePricePage.Go_To_New_Navigate();
        //    servicePricePage.UnfoldAll();
        //    DatasheetDetailsPage datasheetDetailsPage = servicePricePage.EditFirstItem();
        //    servicePricePage.Close();
        //    RecipesVariantPage recipesVariantPage = datasheetDetailsPage.EditRecipe();
        //    //RecipesVariantPage recipesVariantPage = datasheetDetailsPage.EditFirstRecipeItem();
        //    double grossQty1 = recipesVariantPage.GetGrossQty(",");
        //    recipesVariantPage.Close();
        //    qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, serviceNameVacuum);
        //    servicePricePage = qtyAjustementPage.EditFirstService();
        //    servicePricePage.Go_To_New_Navigate();
        //    servicePricePage.UnfoldAll();
        //    datasheetDetailsPage = servicePricePage.EditFirstItem();
        //    servicePricePage.Close();
        //    datasheetDetailsPage.Go_To_New_Navigate();
        //    recipesVariantPage = datasheetDetailsPage.EditRecipe();
        //    //recipesVariantPage = datasheetDetailsPage.EditFirstRecipeItem();
        //    double grossQty2 = recipesVariantPage.GetGrossQty(",");
        //    recipesVariantPage.Close();
        //    prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
        //    prodManagement.ResetFilter();
        //    qtyAjustementPage = prodManagement.DoneToQtyAjustement();
        //    qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Services, "ALL");
        //    double qty1 = qtyAjustementPage.GetQuantity(",", serviceNameNormal);
        //    double qty2 = qtyAjustementPage.GetQuantity(",", serviceNameVacuum);
        //    resultsTab = qtyAjustementPage.GoToResultPage();
        //    string supplyOrderNumber = resultsTab.CreateSupplyOrder(deliveryLocation);
        //    SupplyOrderItem supplyOrderItemPage = resultsTab.GenerateSupplyOrder();
        //    double physQuatityOF = supplyOrderItemPage.GetProdQuantitySum();
        //    List<string> outputFormItemsNames = supplyOrderItemPage.GetItemNames();
        //    Assert.AreEqual(outputFormItemsNames.Count, itemsNames.Count, "the item names are not equal cas 1");
        //    Assert.IsTrue(!outputFormItemsNames.Except(itemsNames).Any(), "the item names are not equal cas 2");
        //    double totQty1 = grossQty1 * qty1;
        //    double totQty2 = grossQty2 * qty2;
        //    double physicalQty = totQty1 + totQty2;
        //    // il manque quelques instructions que ne sont pas tres claire dans la description du test
        //    Assert.AreEqual(physicalQty, physQuatityOF * 10.0, 0.001);
        //}

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Quantity()
        {
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceNameNormal = TestContext.Properties["Prodman_ServiceNormal"].ToString();
            DateTime dateFrom = DateUtils.Now.AddDays(-4);
            DateTime dateTo = DateUtils.Now.AddDays(-1);

            // Arrange 
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.StartTime, "00:00");
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.EndTime, "23:59");
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.Site, siteACE);
            qtyAjustementPage.PageSize("100");
            bool isServiceNormal = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateFrom, dateFrom);
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, dateTo);
            bool isServiceNormalAvant = qtyAjustementPage.IsServicePresent(serviceNameNormal);
            Assert.IsTrue(isServiceNormalAvant, "le service normal n est pas disponible ");
            int Total = Convert.ToInt32(qtyAjustementPage.getTotalService(serviceNameNormal));
            qtyAjustementPage.OpenFirstQtyDetailsPopUp();
            bool isTotalOK = qtyAjustementPage.VerifyQtyTotal(Total);
            Assert.IsTrue(isTotalOK, "Les quantités prévisionnelles doivent être celles qui sont inscrites dans le détail des services à produire dans le vol.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_RenameFavoris()
        {
            HomePage homePage = LogInAsAdmin();

            try
            {
                PageObjects.Parameters.User.ParametersUser parametersUser = homePage.GoToParameters_User();
                parametersUser.SwitchTabs(xPath: ParametersUser.ROLES_TAB);
                parametersUser.ClickProduction();
                parametersUser.ClickProductionInListe();
                parametersUser.ClickProductionManagement();
                parametersUser.CocherRenameFavoritetrenameallfavorites();
                FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
                bool existance = filterAndFavorite.ExistanceFavorite();
                if (existance)
                {
                    filterAndFavorite.ClickEditButton();
                    bool existeinputmodif = filterAndFavorite.EditInput();
                    Assert.IsTrue(existeinputmodif, "Modification Impossible");
                }
                else
                {
                    string nomFavorite = "NomFavorite";
                    string dateTimeNow = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string nomFavoriteAvecDateHeure = $"{nomFavorite}_{dateTimeNow}";
                    filterAndFavorite.MakeFavorite(nomFavoriteAvecDateHeure);
                    filterAndFavorite.SetFavoriteText(nomFavoriteAvecDateHeure);
                    filterAndFavorite.ClickEditButton();
                    bool existeinputmodif = filterAndFavorite.EditInput();
                    Assert.IsTrue(existeinputmodif, "Modification Impossible");
                }
            }
            finally
            {
                homePage.Navigate();
                PageObjects.Parameters.User.ParametersUser parametersUser = homePage.GoToParameters_User();
                parametersUser.SwitchTabs(xPath: ParametersUser.ROLES_TAB);
                parametersUser.ClickProduction();
                parametersUser.ClickProductionInListe();
                parametersUser.ClickProductionManagement();
                parametersUser.UncheckRenameFavoritetrenameallfavorites();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_FilterSubGroupFavori()
        {
            string nomFavorite = "NomFavorite";
            string dateTimeNow = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.UnselectAllItemsFromSubGroup();
            string nomFavoriteAvecDateHeure = $"{nomFavorite}_{dateTimeNow}";
            filterAndFavorite.MakeFavorite(nomFavoriteAvecDateHeure);
            filterAndFavorite.ResetFilter();
            filterAndFavorite.SetFavoriteText(nomFavoriteAvecDateHeure);

            // Assert 1 : La sauvegarde d'un favori 
            bool ResultFilterFavorite = filterAndFavorite.ExistanceFavorite();
            Assert.IsTrue(ResultFilterFavorite, "La sauvegarde du favori a échoué. Assurez-vous qu'aucun sous-groupe n'est sélectionné si cela est requis.");
            ResultPage resultPage = filterAndFavorite.SelectFavoriteName(nomFavoriteAvecDateHeure);

            // Assert 2 :  la non sélection de sous groupes
            bool isAnySubGroupSelected = resultPage.IsthereAnySubGroupSelected();
            Assert.IsTrue(isAnySubGroupSelected, "Le filtre de sub group est encore selectionner ");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_PictoService()
        {
            string pectoVert = "rgb(104, 162, 4)";
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            DateTime dtfrom = DateTime.Today.AddDays(-11);
            DateTime dtto = DateTime.Today;
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, siteACE);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateFrom, dtfrom);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateTo, dtto);
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            var firstserviceName_withdatasheet = qtyAjustementPage.GetFirstServiceNameWithDataSheet();
            var colorservicenamewithdatasheet = qtyAjustementPage.GetClorFirstServiceNameWithDataSheet();
            DatasheetPage datasheetPage = qtyAjustementPage.ClickFirstPictoVert();
            datasheetPage.Go_To_New_Navigate();
            var servicenameinfiltredatasheet = datasheetPage.GetServiceNameInFilterDataSheet();
            var total_number_in_datasheet = datasheetPage.CheckTotalNumber();
            var verifier_name_service = firstserviceName_withdatasheet.Equals(servicenameinfiltredatasheet);
            var picto_vert_service = pectoVert.Equals(colorservicenamewithdatasheet);
            Assert.IsTrue(picto_vert_service, "le picto en vert n'est pas affiché");
            Assert.IsTrue(verifier_name_service && total_number_in_datasheet > 0, "le click sur le picto vert n'amène pas au datasheet");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_FD_NotRounded()
        {

            DateTime dtfrom = DateTime.Today.AddDays(-15);

            // Prepare
            string deliveryLocation = "Produccion";
            string qty = 10.ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateFrom, dtfrom);
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();

            // Génération supply order
            var resultPage = new ResultPage(WebDriver, TestContext);

            resultPage.CreateSupplyOrder(deliveryLocation);
            resultPage.ClickPrefillQuantitiesCheckbox();
            resultPage.GenerateSupplyOrder();

            var itemPage = new ItemPage(WebDriver, TestContext);
            itemPage.ClickOnFirstItemSupplyOrder();

            // Modifier la quantité Prod Qty
            itemPage.SetProdQty(qty);

            // Vérification que F&D n'est pas arrondi
            bool isFDNotRounded = itemPage.VerifyFDNotRounded();

            // Assert
            Assert.IsTrue(isFDNotRounded, "F&D has been rounded to match Prod Qty");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Calcul_SO_Result_Storage()
        {
            DateTime dtfrom = DateTime.Today;
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string deliveryLocation = "Produccion";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateFrom, dtfrom);
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            // Génération supply order
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            string supplyOrderNumber = resultPage.CreateSupplyOrder(deliveryLocation);
            SupplyOrderItem So_Page = resultPage.GenerateSupplyOrder();
           
            List<string>  All_storage = So_Page.GetAllStorage();
            string first_Storage = All_storage[0];

            So_Page.SelectFirstItemNewTab();
            string first_item_name = So_Page.GetFirstItemName();
          
            ItemGeneralInformationPage item_Page = So_Page.EditItemNew(first_item_name);
            So_Page.Close(); 
            item_Page.SearchPackagingProdMan(siteACE);
            string prod_qty = item_Page.GetFirstPackagingProdQtyProdMan();
            string store_qty = item_Page.GetFirstPackagingStorageQtyForProdMan();
            string store_unit = item_Page.GetFirstPackagingStorUnit();
            // 10 KG(10 UD) 
            string storage_from_item = store_qty + " " + store_unit + "(" + prod_qty + ")";
            Assert.AreEqual(storage_from_item, first_Storage, "storage is not same from SO to item interface "); 
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PR_PRODMAN_Calcul_SO_Result_Total_wo_VAT()
        {
            DateTime dtfrom = DateTime.Today;
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string deliveryLocation = "Produccion";
            //Arrange
            HomePage homePage = LogInAsAdmin();
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateFrom, dtfrom);
            QuantityAdjustmentsPage qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            // Génération supply order
            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            string supplyOrderNumber = resultPage.CreateSupplyOrder(deliveryLocation);
            SupplyOrderItem So_Page = resultPage.GenerateSupplyOrder();
            List<double> all_Prices = So_Page.GetAllPrice();
            List<double> all_Qty = So_Page.GetAllProdQty();
            Assert.AreEqual(all_Prices.Count, all_Qty.Count , "The count between price and prod is not the same , verify the parsing ");
            List<Tuple<double, double>> priceQuantityPairs = new List<Tuple<double, double>>();
            for (int i = 0; i < all_Prices.Count; i++)
            {
                priceQuantityPairs.Add(new Tuple<double, double>(all_Prices[i], all_Qty[i]));
            }
            List<double> multiplicationResults = So_Page.GetPriceQuantityMultiplication(priceQuantityPairs);
           double totalVAT = multiplicationResults.Sum();
            double total_VAT_SO = So_Page.GetTotalVat();
            Assert.AreNotEqual(0.0, total_VAT_SO, "Total w/o Vat ");
            Assert.AreEqual(totalVAT, total_VAT_SO , "the total VAT is not same as the calculation");  
        }
        // __________________________________________ Utilitaire ________________________________________________


        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }

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
}