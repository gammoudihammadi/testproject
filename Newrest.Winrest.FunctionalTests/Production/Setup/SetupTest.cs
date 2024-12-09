using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.DeliveryRound;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Setup;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig;
using System.IO.Compression;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Needs;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Catalogs;

namespace Newrest.Winrest.FunctionalTests.Production
{
    [TestClass]
    public class SetupTest : TestBase
    {
        private const int Timeout = 600000;
        string favouriteName = "testSetup" + new Random().Next(10, 99).ToString();
        // données tests
        readonly static string today = DateUtils.Now.ToString("dd/MM/yyyy");
        readonly static DateTime yesterday = DateUtils.Now.AddDays(-1);

        // données pré config (attention données pré-confi aussi dans Production)
        readonly string menu2Name = "Menu PR ACE 2";
        readonly string menu3Name = "Menu PR MAD 1";
        readonly string menu4Name = "Menu PR MAD 2";
        readonly string menu5Name = "Menu PR ACE 3";

        readonly string item2Name = "Item PR ACE 2";
        readonly string item3Name = "Item PR MAD 1";
        readonly string item4Name = "Item PR MAD 2";
        readonly string item5Name = "Item PR ACE 3";

        readonly string priceList1Name = "Price PR ACE";
        readonly string priceList2Name = "Price PR MAD";

        readonly string delivery2Name = "Delivery PR ACE 2";
        readonly string delivery3Name = "Delivery PR MAD 1";
        readonly string delivery4Name = "Delivery PR MAD 2";

        readonly string delivery5Name = "Delivery PR ACE 3";

        readonly string service2Name = "Service PR ACE 2";
        readonly string service3Name = "Service PR MAD 1";
        readonly string service4Name = "Service PR MAD 2";
        readonly string service5Name = "Service PR ACE 3";

        readonly string deliveryRound2Name = "DR PR ACE 2";
        readonly string deliveryRound3Name = "DR PR MAD 1";
        readonly string deliveryRound4Name = "DR PR MAD 2";

        readonly string customer2Name = "Customer PR ACE 2";
        readonly string customer3Name = "Customer PR MAD 1";
        readonly string customer4Name = "Customer PR MAD 2";
        readonly string customer5Name = "Customer PR ACE 3";

        readonly string customer2Code = "PC2";
        readonly string customer3Code = "PC3";
        readonly string customer4Code = "PC4";
        readonly string customer5Code = "PC5";

        readonly string recipe2Name = "Recipe PR ACE 2";
        readonly string recipe3Name = "Recipe PR MAD 1";
        readonly string recipe4Name = "Recipe PR MAD 2";
        readonly string recipe5Name = "Recipe PR ACE 3";

        readonly string itemTest = "ITEM Test PR";

        //config data v2

        //foodpack
        readonly string foodpackIndividual = "Individual";
        readonly string foodpackMultiporcion = "Multiporcion";
        readonly string foodpackB1 = "B1";
        readonly string foodpackB4 = "B4";
        readonly string foodpackB6 = "B6";
        readonly string foodpackBACGN = "BAC G/N";
        readonly string foodpackBACGNhalf = "BAC G/N 1/2";
        readonly string foodpackDemi = "DEMI";
        readonly string foodpackEntier = "ENTIER";
        readonly string foodpackUnit = "UNIT";
        readonly string foodpackCRT = "CRT";
        readonly string foodpackSEAU = "SEAU";

        readonly string portionsFoodpackB1 = "1";
        readonly string portionsFoodpackB4 = "4";
        readonly string portionsFoodpackB6 = "6";
        readonly string portionsFoodpackBACGN = "10";
        readonly string portionsFoodpackBACGNhalf = "5";

        //recipe type
        readonly string recipeTypeNP = "RECETA NP";
        readonly string recipeTypeBulk = "RECETA BULK";

        //Data Delivery ES (3 services, packaging Individual/Multiporcion)
        readonly string customerESName = "ES CUSTOMER";
        readonly string customerESCode = "ESC";

        readonly string ESDeliveryName = "ES DELIVERY";
        readonly string packagingMethodIndividualMultiporcion = "Individual / Multi-Portion";

        readonly string ESINDServicioName = "ES IND SERVICIO";
        readonly string ESINDMenuName = "ES IND MENU";
        readonly string ESINDCoefficient = "50";
        readonly string ESINDServicioPaxMin = "1";
        readonly string ESINDServicioPaxMax = "9999";
        readonly string ESINDServicioPortionsNumber = "1";

        readonly string ESMULServicioName = "ES MUL SERVICIO";
        readonly string ESMULMenuName = "ES MUL MENU";
        readonly string ESMULCoefficient = "100";
        readonly string ESMULServicioPaxMin = "1";
        readonly string ESMULServicioPaxMax = "9999";
        readonly string ESMULServicioPortionsNumber = "9999";

        readonly string ESINDMULServicioName = "ES IND MUL SERVICIO";
        readonly string ESINDMULMenuName = "ES IND MUL MENU";
        readonly string ESINDMULCoefficient = "150";
        readonly string ESINDMULIndividualServicioPaxMin = "1";
        readonly string ESINDMULIndividualServicioPaxMax = "4";
        readonly string ESINDMULIndividualServicioPortionsNumber = "1";
        readonly string ESINDMULMultiporcionServicioPaxMin = "5";
        readonly string ESINDMULMultiporcionServicioPaxMax = "9999";
        readonly string ESINDMULMultiporcionServicioPortionsNumber = "9999";

        readonly string recetaESName = "ES RECETA";
        readonly string portion150g = "150";
        readonly string deliveryRoundES = "DR ES";

        // Datas Delivery TLS (packaging pax per packaging, 3 foodpacks)
        readonly string TLSCustomerName = "TLS CUSTOMER";
        readonly string TLSCustomerCode = "TLSC";
        readonly string TLSServicioName = "TLS SERVICIO";
        readonly string TLSRecetaName = "RECETA TLS";
        readonly string TLSMenuName = "MENU TLS";
        readonly string TLSCoefficient200 = "200";
        readonly string TLSDeliveryName = "DELIVERY TLS";
        readonly string methodPaxPerPackaging = "PAX per packaging";
        readonly string TLSDeliveryRound = "TLS DR";

        readonly string customerTLSGRName = "CUSTOMER GR TLS";
        readonly string customerTLSGRCode = "TL GR";
        readonly string portion100g = "100";
        readonly string servicioTLSGR1Name = "SERVICIO GR TLS 1";
        readonly string servicioTLSGR2Name = "SERVICIO GR TLS 2";
        readonly string recetaTLSGRName = "RECETA GR TLS";
        readonly string menuTLSGR1Name = "MENU GR TLS 1";
        readonly string menuTLSGR2Name = "MENU GR TLS 2";
        readonly string deliveryTLSGRName = "DELIVERY GR TLS";
        readonly string deliveryRoundTLSGR = "DR GR TLS";
        readonly string TLSGRServicio1DispatchQuantity = "5";
        readonly string TLSGRServicio2DispatchQuantity = "10";

        readonly string methodByRecipeVariant = "By recipe-variant";
        readonly string customerPFName = "CUSTOMER PF";
        readonly string customerPFCode = "PF";
        readonly string servicioPFName = "SERVICIO PF";
        readonly string recetaPFName = "RECETA PF";
        readonly string menuPFName = "MENU PF";
        readonly string yieldPF95 = "95";
        readonly string deliveryPFName = "DELIVERY PF";
        readonly string deliveryRoundPF = "DEL RD PF";

        readonly string condrenCustomerName = "CUSTOMER CONDREN";
        readonly string condrenCustomerCode = "COND";
        readonly string condrenServicioName = "SERVICIO CONDREN";
        readonly string condrenReceta1Name = "RECETA CONDREN 1";
        readonly string condrenReceta2Name = "RECETA CONDREN 2";
        readonly string condrenMenuName = "MENU CONDREN";
        readonly string condrenDeliveryName = "DELIVERY CONDREN";
        readonly string condrenDeliveryRound = "DR CONDREN";

        readonly string barentinCustomerName = "CUSTOMER BARENTIN";
        readonly string barentinCustomerCode = "BAR";
        readonly string barentinServicioAdulto = "SERVICIO BARENTIN ADULTO";
        readonly string barentinServicioCollegio = "SERVICIO BARENTIN COLLEGIO";
        readonly string barentinReceta1 = "RECETA BARENTIN 1";
        readonly string barentinRecetaNP2 = "RECETA BARENTIN NP 2";
        readonly string barentinRecetaBulk3 = "RECETA BARENTIN BULK 3";
        readonly string barentinMenuAdulto = "MENU BARENTIN ADULTO";
        readonly string barentinMenuCollegio = "MENU COLLEGIO";
        readonly string barentinDeliveryName = "DELIVERY BARENTIN";
        readonly string barentinDeliveryRound = "DR BAR";
        readonly string methodGroupedByRecipe = "Grouped by recipe";
        readonly string CUSTOMERTYPE = "Colectividades";


        [TestInitialize]
        public override void TestInitialize()
        {

            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(PR_SETUP_Filter_ByDeliveryRoundName):
                    PR_SETUP_Create_Delivery_Rounds_TestInitialize();
                    break;
                case nameof(PR_SETUP_Filter_ByDate):
                    PR_SETUP_Create_Delivery_Rounds_TestInitialize();
                    break;
                case nameof(PR_SETUP_Filter_ByRecipeName):
                    PR_SETUP_Create_Recipes_TestInitialize();
                    break;
                case nameof(PR_SETUP_Filter_ByCustomers):
                    PR_SETUP_Create_Customers_TestInitialize();
                    break;
                case nameof(PR_SETUP_Filter_ByDeliveries):
                    PR_SETUP_Create_Deliveries_TestInitialize();
                    break;
                case nameof(PR_SETUP_Link_Recipe):
                    PR_SETUP_Create_Delivery_Rounds_TestInitialize();
                    break;
                case nameof(PR_SETUP_ExportCSVFile):
                    PR_SETUP_Create_Delivery_Rounds_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckPAX_MethodIndividualMultiportion):
                    PR_SETUP_Create_Datas_MethodIndividualMultiportion_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckPAX_MethodPaxPerPackaging):
                    PR_SETUP_Create_Datas_MethodPaxPerPackaging_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckPAX_MethodPaxPerPackagingGrouped):
                    PR_SETUP_Create_Datas_MethodPaxPerPackagingGrouped_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckPAX_MethodByRecipeVariant):
                    PR_SETUP_Create_Datas_MethodByRecipeVariant_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckPAX_MethodByRecipeVariant2Recipes):
                    PR_SETUP_Create_Datas_MethodByRecipeVariant2Recipes_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckPAX_MethodRecipeVariantAndGroupedByRecipe):
                    PR_SETUP_Create_Datas_MethodRecipeVariantAndGroupedByRecipe_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckQuantityToProduce_MethodIndividualMultiportion):
                    PR_SETUP_Create_Datas_MethodIndividualMultiportion_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckQuantityToProduce_MethodPaxPerPackaging):
                    PR_SETUP_Create_Datas_MethodPaxPerPackaging_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckQuantityToProduce_MethodPaxPerPackagingGrouped):
                    PR_SETUP_Create_Datas_MethodPaxPerPackagingGrouped_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckQuantityToProduce_MethodByRecipeVariant):
                    PR_SETUP_Create_Datas_MethodByRecipeVariant_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckQuantityToProduce_MethodByRecipeVariant2Recipes):
                    PR_SETUP_Create_Datas_MethodByRecipeVariant2Recipes_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckQuantityToProduce_MethodRecipeVariantAndGroupedByRecipe):
                    PR_SETUP_Create_Datas_MethodRecipeVariantAndGroupedByRecipe_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckWeightPerPack_MethodIndividualMultiportion):
                    PR_SETUP_Create_Datas_MethodIndividualMultiportion_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckWeightPerPack_MethodPaxPerPackaging):
                    PR_SETUP_Create_Datas_MethodPaxPerPackaging_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckWeightPerPack_MethodPaxPerPackagingGrouped):
                    PR_SETUP_Create_Datas_MethodPaxPerPackagingGrouped_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckWeightPerPack_MethodByRecipeVariant):
                    PR_SETUP_Create_Datas_MethodByRecipeVariant_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckWeightPerPack_MethodByRecipeVariant2Recipes):
                    PR_SETUP_Create_Datas_MethodByRecipeVariant2Recipes_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckWeightPerPack_MethodRecipeVariantAndGroupedByRecipe):
                    PR_SETUP_Create_Datas_MethodRecipeVariantAndGroupedByRecipe_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckTotal_MethodIndividualMultiportion):
                    PR_SETUP_Create_Datas_MethodIndividualMultiportion_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckTotal_MethodPaxPerPackaging):
                    PR_SETUP_Create_Datas_MethodPaxPerPackaging_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckTotal_MethodPaxPerPackagingGrouped):
                    PR_SETUP_Create_Datas_MethodPaxPerPackagingGrouped_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckTotal_MethodByRecipeVariant):
                    PR_SETUP_Create_Datas_MethodByRecipeVariant_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckTotal_MethodByRecipeVariant2Recipes):
                    PR_SETUP_Create_Datas_MethodByRecipeVariant2Recipes_TestInitialize();
                    break;
                case nameof(PR_SETUP_Display_CheckTotal_MethodRecipeVariantAndGroupedByRecipe):
                    PR_SETUP_Create_Datas_MethodRecipeVariantAndGroupedByRecipe_TestInitialize();
                    break;
                case nameof(PR_SETUP_Export_CheckNbLabel):
                    PR_SETUP_SetProductionSettings_TestInitialize();
                    PR_SETUP_Create_Datas_MethodPaxPerPackaging_TestInitialize();
                    break;
                case nameof(PR_SETUP_Export_CheckSiteName):
                    PR_SETUP_SetProductionSettings_TestInitialize();
                    PR_SETUP_Create_Datas_MethodIndividualMultiportion_TestInitialize();
                    break;
                case nameof(PR_SETUP_Export_CheckMenuName):
                    PR_SETUP_Create_Datas_MethodIndividualMultiportion_TestInitialize();
                    PR_SETUP_Create_Datas_MethodPaxPerPackaging_TestInitialize();
                    PR_SETUP_Create_Datas_MethodPaxPerPackagingGrouped_TestInitialize();
                    break;
                case nameof(PR_SETUP_SelectFavorite):
                    PR_SETUP_Create_Favorite_TestInitialize();
                    break;
                case nameof(PR_SETUP_FilterWorkshop):
                    PR_SETUP_Create_Datas_MethodByRecipeVariant2Recipes_TestInitialize();
                    break;

                case nameof(PR_SETUP_FAF_DONE):
                    PR_SETUP_Create_Customers_TestInitialize();
                    PR_SETUP_Create_Services_TestInitialize();
                    PR_SETUP_Create_Deliveries_TestInitialize();
                    PR_SETUP_Create_Delivery_Rounds_TestInitialize();

                    break;
                case nameof(PR_SETUP_SearchByDeliveryRound):
                    PR_SETUP_Create_Customers_TestInitialize();
                    PR_SETUP_Create_Services_TestInitialize();
                    PR_SETUP_Create_Deliveries_TestInitialize();
                    PR_SETUP_Create_Delivery_Rounds_TestInitialize();
                    PR_SETUP_ValidateDispatchs_TestInitialize();
                    break;
                case nameof(PR_SETUP_FilterDate):
                    PR_SETUP_Create_Customers_TestInitialize();
                    PR_SETUP_Create_Services_TestInitialize();
                    PR_SETUP_Create_Deliveries_TestInitialize();
                    PR_SETUP_Create_Delivery_Rounds_TestInitialize();
                    PR_SETUP_ValidateDispatchs_TestInitialize();
                    break;
                case nameof(PR_SETUP_SearchByRecipeName):
                    PR_SETUP_Create_Datas_MethodByRecipeVariant2Recipes_TestInitialize();
                    break;
                case nameof(PR_SETUP_FAF_GoToResults):
                    PR_SETUP_Create_Customers_TestInitialize();
                    PR_SETUP_Create_Services_TestInitialize();
                    PR_SETUP_Create_Deliveries_TestInitialize();
                    break;

                case nameof(PR_SETUP_FilterCustomer):
                    PR_SETUP_Create_Customers_TestInitialize();
                    PR_SETUP_Create_Services_TestInitialize();
                    PR_SETUP_Create_Deliveries_TestInitialize();
                    PR_SETUP_Create_Delivery_Rounds_TestInitialize();
                    PR_SETUP_ValidateDispatchs_TestInitialize();
                    break;
                case nameof(PR_SETUP_FilterCustomerType):
                    PR_SETUP_Create_Datas_MethodByRecipeVariant2Recipes_TestInitialize();
                    break;
                default:
                    break;
            }
        }

        //[TestCleanup]
        //public override void TestCleanup()
        //{
        //    var testMethod = TestContext.TestName;
        //    switch (testMethod)
        //    {
        //        case nameof(PR_SETUP_SelectFavorite):
        //            PR_SETUP_SelectFavorite_TestCleanup();
        //            break;
        //        default:
        //            break;
        //    }
        //}
        public void PR_SETUP_Create_Favorite_TestInitialize()
        {
            // Prepare    
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);

            //Create Favorite
            setupFilterAndFavoritesPage.MakeFavorite(favouriteName);
            setupFilterAndFavoritesPage.SetFavoriteText(favouriteName);

            //Assert
            Assert.IsTrue(setupFilterAndFavoritesPage.IsFavoritePresent(favouriteName), "Le favori " + favouriteName + " n'a pas été créé.");
            setupFilterAndFavoritesPage.ResetFilters();
        }

        //private void PR_SETUP_SelectFavorite_TestCleanup()
        //{
        //    //arrange
        //    //Prepare
        //    string siteACE = TestContext.Properties["Production_Site1"].ToString();

        //    //Arrange
        //    HomePage homePage = LogInAsAdmin();

        //    var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();

        //    //Filter 
        //    setupFilterAndFavoritesPage.ResetFilters();
        //    setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
        //    setupFilterAndFavoritesPage.SetFavoriteText(favouriteName);
        //    setupFilterAndFavoritesPage.DeleteFavorite(favouriteName);
        //}
        public void PR_SETUP_SetProductionSettings_TestInitialize()
        {
            //Variants
            string guestTypeVariant1 = TestContext.Properties["Production_GuestType1"].ToString();
            string guestTypeVariant2 = TestContext.Properties["Production_GuestType2"].ToString();
            string mealVariant1 = TestContext.Properties["Production_Meal1"].ToString();
            string mealVariant2 = TestContext.Properties["Production_Meal2"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            ClearCache();
            ParametersProduction parametersProduction = homePage.GoToParameters_ProductionPage();
            parametersProduction.GoToTab_Variant();
            // Variant 1
            parametersProduction.FilterSite(siteACE);
            parametersProduction.FilterGuestType(guestTypeVariant1);
            parametersProduction.FilterMealType(mealVariant1);

            if (!parametersProduction.IsVariantPresent(guestTypeVariant1, mealVariant1))
            {
                parametersProduction.AddNewVariant(guestTypeVariant1, mealVariant1, siteACE);
            }

            parametersProduction.FilterSite(siteACE);
            parametersProduction.FilterGuestType(guestTypeVariant1);
            parametersProduction.FilterMealType(mealVariant1);

            Assert.IsTrue(parametersProduction.IsVariantPresent(guestTypeVariant1, mealVariant1));

            // Variant 2
            parametersProduction.FilterSite(siteMAD);
            parametersProduction.FilterGuestType(guestTypeVariant2);
            parametersProduction.FilterMealType(mealVariant2);

            if (!parametersProduction.IsVariantPresent(guestTypeVariant2, mealVariant2))
            {
                parametersProduction.AddNewVariant(guestTypeVariant2, mealVariant2, siteMAD);
            }

            parametersProduction.FilterSite(siteMAD);
            parametersProduction.FilterGuestType(guestTypeVariant2);
            parametersProduction.FilterMealType(mealVariant2);

            Assert.IsTrue(parametersProduction.IsVariantPresent(guestTypeVariant2, mealVariant2));

            //Create FoodPacks
            parametersProduction.GoToTab_Foodpack();

            if (!parametersProduction.CheckFoodPackExists(foodpackIndividual))
            {
                parametersProduction.CreateNewFoodPack(foodpackIndividual);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackMultiporcion))
            {
                parametersProduction.CreateNewFoodPack(foodpackMultiporcion);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackB1))
            {
                parametersProduction.CreateNewFoodPack(foodpackB1);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackB4))
            {
                parametersProduction.CreateNewFoodPack(foodpackB4);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackB6))
            {
                parametersProduction.CreateNewFoodPack(foodpackB6);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackBACGN))
            {
                parametersProduction.CreateNewFoodPack(foodpackBACGN);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackBACGNhalf))
            {
                parametersProduction.CreateNewFoodPack(foodpackBACGNhalf);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackDemi))
            {
                parametersProduction.CreateNewFoodPack(foodpackDemi);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackEntier))
            {
                //Add Bulk Equivalence to foodpack
                parametersProduction.CreateNewFoodPack(foodpackEntier);
                parametersProduction.AddBulkEquivalenceToFoodPack(foodpackEntier, foodpackDemi, "2");
                var bulkEquivalence = parametersProduction.GetBulkEquivalence(foodpackEntier);
                Assert.IsTrue(new[] { foodpackDemi, "3" }.Any(eq => bulkEquivalence.Contains(eq)), "L'équivalence pour le foodpack {0} n'est pas correcte : {1} au lieu de {2} x {3}.", foodpackEntier, bulkEquivalence, "2", foodpackDemi);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackUnit))
            {
                parametersProduction.CreateNewFoodPack(foodpackUnit);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackCRT))
            {
                parametersProduction.CreateNewFoodPack(foodpackCRT);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackSEAU))
            {
                parametersProduction.CreateNewFoodPack(foodpackSEAU);
            }

            //Create Recipe Type
            parametersProduction.GoToTab_RecipeType();
            if (!parametersProduction.CheckRecipeTypeExists(recipeTypeNP))
            {
                parametersProduction.CreateNewRecipeType(recipeTypeNP, "16");
            }

            if (!parametersProduction.CheckRecipeTypeExists(recipeTypeBulk))
            {
                parametersProduction.CreateNewRecipeType(recipeTypeBulk, "17");
            }
        }

        public void PR_SETUP_Create_Customers_TestInitialize()
        {
            //warn : dependances deliveries filtres
            // Prepare            
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string customerType2 = TestContext.Properties["Production_CustomerType2"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            //second customer for Production
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer2Name);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer2Name, customer2Code, customerType2);
                CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer2Name);
            }
            string firstCustomerName = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customer2Name, firstCustomerName, "Le customer n'a pas été créé.");
            //third customer for Production
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer3Name);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer3Name, customer3Code, customerType1);
                CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer3Name);
            }
            firstCustomerName = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customer3Name, firstCustomerName, "Le customer n'a pas été créé.");
            //fourth customer for Production
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer4Name);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer4Name, customer4Code, customerType2);
                CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer4Name);
            }
            firstCustomerName = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customer4Name, firstCustomerName, "Le customer n'a pas été créé.");
            //fifth customer for Production (no delivery round)
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer5Name);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer5Name, customer5Code, customerType2);
                CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer5Name);
            }
            firstCustomerName = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customer5Name, firstCustomerName, "Le customer n'a pas été créé.");
        }

        public void PR_SETUP_Create_Services_TestInitialize()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string customer2Reference = $"{customer2Code} - {customer2Name}";
            string customer3Reference = $"{customer3Code} - {customer3Name}";
            string customer4Reference = $"{customer4Code} - {customer4Name}";
            string customer5Reference = $"{customer5Code} - {customer5Name}";
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            // service 2
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service2Name);
            if (servicePage.CheckTotalNumber() == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(service2Name, null, null, serviceCategorie, null, serviceType);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                serviceGeneralInformationsPage.SetProduced(true);
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customer2Reference, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                ServicePricePage pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                ServiceCreatePriceModalPage serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customer2Name);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }
                ServiceGeneralInformationPage serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service2Name);
            string service2 = servicePage.GetFirstServiceName();
            Assert.IsTrue(service2.Contains(service2Name), $"Le service {service2Name} n'existe pas.");

            // service 3
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service3Name);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(service3Name, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteMAD, customer3Reference, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteMAD, customer3Name);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service3Name);

            string service3 = servicePage.GetFirstServiceName();

            Assert.IsTrue(service3.Contains(service3Name), "Le service " + service3Name + " n'existe pas.");

            // service 4
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service4Name);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(service4Name, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteMAD, customer4Reference, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteMAD, customer4Name);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service4Name);

            string service4 = servicePage.GetFirstServiceName();

            Assert.IsTrue(service4.Contains(service4Name), "Le service " + service4Name + " n'existe pas.");

            // service 5 no delivery round
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service5Name);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(service5Name, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customer5Reference, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customer5Name);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service5Name);

            string service5 = servicePage.GetFirstServiceName();

            Assert.IsTrue(service5.Contains(service5Name), "Le service " + service5Name + " n'existe pas.");
        }

        public void PR_SETUP_Create_Items_TestInitialize()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            string group1 = TestContext.Properties["Production_ItemGroup1"].ToString();
            string group2 = TestContext.Properties["Production_ItemGroup2"].ToString();

            string workshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string workshop2 = TestContext.Properties["Production_Workshop2"].ToString();

            string supplier1 = TestContext.Properties["Production_ItemSupplier1"].ToString();
            string supplier2 = TestContext.Properties["Production_ItemSupplier2"].ToString();

            string taxType = TestContext.Properties["Production_ItemTaxType"].ToString();
            string prodUnit = TestContext.Properties["Production_ItemProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();

            // ----------------------------------------- Item 2 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, item2Name);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item2Name, group2, workshop2, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier2);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, item2Name);
            }

            string item2 = itemPage.GetFirstItemName();

            // ----------------------------------------- Item 3 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, item3Name);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item3Name, group1, workshop1, taxType, prodUnit);

                // 1 packaging pour le site MAD
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteMAD, packagingName, storageQty, storageUnit, qty, supplier1);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, item3Name);
            }

            string item3 = itemPage.GetFirstItemName();

            // ----------------------------------------- Item 2 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, item4Name);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item4Name, group2, workshop2, taxType, prodUnit);

                // 1 packaging pour le site MAD
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteMAD, packagingName, storageQty, storageUnit, qty, supplier2);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, item4Name);
            }

            string item4 = itemPage.GetFirstItemName();

            // ----------------------------------------- Item 3 ACE NO DELIVERY ROUND ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, item5Name);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item5Name, group2, workshop2, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier2);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, item5Name);
            }

            string item5 = itemPage.GetFirstItemName();

            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemTest);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemTest, group1, workshop1, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier1);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemTest);
            }

            string item6 = itemPage.GetFirstItemName();

            // Assert
            Assert.AreEqual(item2Name, item2, "L'item " + item2Name + " n'existe pas.");
            Assert.AreEqual(item3Name, item3, "L'item " + item3Name + " n'existe pas.");
            Assert.AreEqual(item4Name, item4, "L'item " + item4Name + " n'existe pas.");
            Assert.AreEqual(item5Name, item5, "L'item " + item5Name + " n'existe pas.");
            Assert.AreEqual(itemTest, item6, "L'item " + itemTest + " n'existe pas.");
        }

        public void PR_SETUP_AddItemsToPriceLists_TestInitialize()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            PageObjects.Customers.PriceList.PriceListPage priceListPage = homePage.GoToCustomers_PriceListPage();
            // price list 1 for ACE
            priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Site, siteACE);
            priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Search, priceList1Name);
            if (priceListPage.CheckTotalNumber() == 0)
            {
                //création de la price list
                PageObjects.Customers.PriceList.PricingCreateModalpage pricingCreateModalpage = priceListPage.CreateNewPricing();
                PageObjects.Customers.PriceList.PriceListDetailsPage priceListDetailsPage = pricingCreateModalpage.FillField_CreateNewPricing(priceList1Name, siteACE, DateUtils.Now, DateUtils.Now, true);
                priceListDetailsPage.AddItem(item2Name, false);
                priceListDetailsPage.AddItem(item5Name, false);
                priceListDetailsPage.BackToList();
                priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Search, priceList1Name);
            }
            int totalNumber = priceListPage.CheckTotalNumber();
            Assert.IsTrue(totalNumber == 1, $"La priceList : {priceList1Name} n'a pas été crée");
            if (totalNumber == 1)
            {
                //La priceList est crée, on vérifié la présence des items
                PageObjects.Customers.PriceList.PriceListDetailsPage priceListDetailsPage = priceListPage.ClickOnFirstPricing();
                if (!priceListDetailsPage.IsItemAdded(item2Name)) priceListDetailsPage.AddItem(item2Name, false);
                if (!priceListDetailsPage.IsItemAdded(item5Name)) priceListDetailsPage.AddItem(item5Name, false);
                //Asserts
                bool isItemAdded = priceListDetailsPage.IsItemAdded(item2Name);
                Assert.IsTrue(isItemAdded, $"L'item : {item2Name} n'a pas été ajouté à la PriceList : {priceList1Name}");
                isItemAdded = priceListDetailsPage.IsItemAdded(item5Name);
                Assert.IsTrue(isItemAdded, $"L'item : {item5Name} n'a pas été ajouté à la PriceList : {priceList1Name}");
                priceListDetailsPage.BackToList();
            }
            priceListPage.ResetFilter();
            // price list for MAD
            priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Site, siteMAD);
            priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Search, priceList2Name);
            totalNumber = priceListPage.CheckTotalNumber();
            if (totalNumber == 0)
            {
                //création de la price list
                PageObjects.Customers.PriceList.PricingCreateModalpage pricingCreateModalpage = priceListPage.CreateNewPricing();
                PageObjects.Customers.PriceList.PriceListDetailsPage priceListDetailsPage = pricingCreateModalpage.FillField_CreateNewPricing(priceList2Name, siteMAD, DateUtils.Now, DateUtils.Now, true);
                priceListDetailsPage.AddItem(item3Name, false);
                priceListDetailsPage.AddItem(item4Name, false);
                priceListDetailsPage.BackToList();
                priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Search, priceList2Name);
            }
            Assert.IsTrue(totalNumber == 1, $"La priceList : {priceList2Name} n'a pas été crée");
            if (totalNumber == 1)
            {
                //La priceList est crée, on vérifié la présence des items
                PageObjects.Customers.PriceList.PriceListDetailsPage priceListDetailsPage = priceListPage.ClickOnFirstPricing();
                if (!priceListDetailsPage.IsItemAdded(item3Name)) priceListDetailsPage.AddItem(item3Name, false);
                if (!priceListDetailsPage.IsItemAdded(item4Name)) priceListDetailsPage.AddItem(item4Name, false);
                //Asserts
                bool isItemAdded = priceListDetailsPage.IsItemAdded(item3Name);
                Assert.IsTrue(isItemAdded, $"L'item : {item3Name} n'a pas été ajouté à la PriceList : {priceList2Name}");
                isItemAdded = priceListDetailsPage.IsItemAdded(item4Name);
                Assert.IsTrue(isItemAdded, $"L'item : {item4Name} n'a pas été ajouté à la PriceList : {priceList2Name}");
            }
        }

        public void PR_SETUP_Create_Recipes_TestInitialize()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeVariantForMAD = TestContext.Properties["Production_RecipeVariant1ForMAD"].ToString();

            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";

            string recipeType2 = TestContext.Properties["Production_RecipeType2"].ToString();
            string recipeCookingMode2 = TestContext.Properties["Production_CookingMode2"].ToString();
            string recipeWorkshop2 = TestContext.Properties["Production_Workshop2"].ToString();
            string recipePortion2 = "18";

            string newTotalPortion = "100";

            //Arrange           
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();

            // ---------------------------------------------- recette 2 --------------------------------------------------

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe2Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipe2Name, recipeType2, recipePortion2, true, null, recipeCookingMode2, recipeWorkshop2);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(item2Name);
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(item2Name);
                }
                recipeVariantPage.SetTotalPortion(newTotalPortion);
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe2Name);

            string recipe2 = recipesPage.GetFirstRecipeName();

            // ---------------------------------------------- recette 3 --------------------------------------------------

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe3Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipe3Name, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteMAD, recipeVariantForMAD);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(item3Name);
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(item3Name);
                }
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe3Name);

            string recipe3 = recipesPage.GetFirstRecipeName();

            // ---------------------------------------------- recette 4 --------------------------------------------------

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe4Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipe4Name, recipeType2, recipePortion2, true, null, recipeCookingMode2, recipeWorkshop2);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteMAD, recipeVariantForMAD);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(item4Name);

                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(item4Name);
                }
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe4Name);

            string recipe4 = recipesPage.GetFirstRecipeName();

            // ---------------------------------------------- recette 5 NO DELIVERY ROUND --------------------------------------------------

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe5Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipe5Name, recipeType2, recipePortion2, true, null, recipeCookingMode2, recipeWorkshop2);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(item5Name);
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(item5Name);
                }
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe5Name);

            string recipe5 = recipesPage.GetFirstRecipeName();

            // --------------------------------------------- Assert ------------------------------------------------------------
            Assert.AreEqual(recipe2Name, recipe2, "La recette " + recipe2Name + " n'existe pas.");
            Assert.AreEqual(recipe3Name, recipe3, "La recette " + recipe3Name + " n'existe pas.");
            Assert.AreEqual(recipe4Name, recipe4, "La recette " + recipe4Name + " n'existe pas.");
            Assert.AreEqual(recipe5Name, recipe5, "La recette " + recipe5Name + " n'existe pas.");
        }

        public void PR_SETUP_Create_Menus_TestInitialize()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            string recipeVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();
            string recipeVariantForMAD = TestContext.Properties["Production_MenuVariant1ForMAD"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();

            // menu 2
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu2Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(service2Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menu2Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, recipeVariantForACE);
                menuDayViewPage.AddRecipe(recipe2Name);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe2Name);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe2Name);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu2Name);
            Assert.AreEqual(menu2Name, menusPage.GetFirstMenuName(), "Le menu2 n'a pas été créé.");


            // menu 3
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu3Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteMAD);
                menusCreateModalPage.AddService(service3Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menu3Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteMAD, recipeVariantForMAD);
                menuDayViewPage.AddRecipe(recipe3Name);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe3Name);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe3Name);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu3Name);
            Assert.AreEqual(menu3Name, menusPage.GetFirstMenuName(), "Le menu n'a pas été créé.");

            // menu 4
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu4Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteMAD);
                menusCreateModalPage.AddService(service4Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menu4Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteMAD, recipeVariantForMAD);
                menuDayViewPage.AddRecipe(recipe4Name);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe4Name);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe4Name);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu4Name);
            Assert.AreEqual(menu4Name, menusPage.GetFirstMenuName(), "Le menu n'a pas été créé.");

            // menu 5 NO DELIVERY ROUND
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu5Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(service5Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menu5Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, recipeVariantForACE);
                menuDayViewPage.AddRecipe(recipe5Name);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe5Name);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe5Name);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu5Name);
            Assert.AreEqual(menu5Name, menusPage.GetFirstMenuName(), "Le menu n'a pas été créé.");
        }

        public void PR_SETUP_Create_Deliveries_TestInitialize()
        {
            // Prepare
            string qty = "10";
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            string packagingMethod1 = "By recipe-variant";
            string packagingName1 = "Individual";
            string packagingMethod2 = "PAX per packaging";
            string packagingName2 = "Multiporcion";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            //delivery 2
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery2Name);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryPage.WaitPageLoading();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(delivery2Name, customer2Name, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                deliveryLoadingPage.AddService(service2Name);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethod2, packagingName2, true);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery2Name);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Equals(delivery2Name), "La delivery : " + delivery2Name + " n'a pas été crée");

            deliveryPage.ResetFilter();

            //delivery 3
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery3Name);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(delivery3Name, customer3Name, siteMAD, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                deliveryLoadingPage.AddService(service3Name);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethod1, packagingName2, false);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery3Name);

            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Equals(delivery3Name), "La delivery : " + delivery3Name + " n'a pas été crée");

            deliveryPage.ResetFilter();

            //delivery 4
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery4Name);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(delivery4Name, customer4Name, siteMAD, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                deliveryLoadingPage.AddService(service4Name);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethod2, packagingName1, true);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery4Name);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Equals(delivery4Name), "La delivery : " + delivery4Name + " n'a pas été crée");

            //delivery 5 NO DELIVERY ROUND
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery5Name);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(delivery5Name, customer5Name, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                deliveryLoadingPage.AddService(service5Name);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethod2, packagingName2, true);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery5Name);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Equals(delivery5Name), "La delivery : " + delivery5Name + " n'a pas été crée");
        }

        public void PR_SETUP_Create_Delivery_Rounds_TestInitialize()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();
            ParametersSites sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteMAD);
            string siteIdMAD = sitePage.CollectNewSiteID();

            // Act

            // delivery round 2
            DeliveryRoundPage deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound2Name);
            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                DeliveryRoundCreateModalpage deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRound2Name, "ACE - ACE", DateUtils.Now, DateUtils.Now.AddDays(+31));
                DeliveryRoundDeliveryPage deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(delivery2Name);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound2Name);
            }
            else
            {
                DeliveryRoundDeliveryPage deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }
            //Assert
            string firstDeliveryRound = deliveryRoundPage.GetFirstDeliveryRound();
            Assert.AreEqual(deliveryRound2Name, firstDeliveryRound, "Le delivery round n'a pas été créé.");
            // delivery round 3
            deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound3Name);
            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                DeliveryRoundCreateModalpage deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRound3Name, "MAD - MAD", DateUtils.Now, DateUtils.Now.AddDays(+31));
                DeliveryRoundDeliveryPage deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(delivery3Name);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound3Name);
            }
            else
            {
                DeliveryRoundDeliveryPage deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }
            //Assert
            firstDeliveryRound = deliveryRoundPage.GetFirstDeliveryRound();
            Assert.AreEqual(deliveryRound3Name, firstDeliveryRound, "Le delivery round n'a pas été créé.");
            // delivery round 4
            deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound4Name);
            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                DeliveryRoundCreateModalpage deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRound4Name, "MAD - MAD", DateUtils.Now, DateUtils.Now.AddDays(+31));
                DeliveryRoundDeliveryPage deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(delivery4Name);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound4Name);
            }
            else
            {
                DeliveryRoundDeliveryPage deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }
            firstDeliveryRound = deliveryRoundPage.GetFirstDeliveryRound();
            //Assert
            Assert.AreEqual(deliveryRound4Name, firstDeliveryRound, "Le delivery round n'a pas été créé.");
        }

        public void PR_SETUP_ValidateDispatchs_TestInitialize()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // dispatch 2
            DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);
            dispatchPage.Filter(DispatchPage.FilterType.Search, delivery2Name);
            PrevisionalQtyPage previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
            // dispatch 3
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteMAD);
            dispatchPage.Filter(DispatchPage.FilterType.Search, delivery3Name);
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
            // dispatch 4
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteMAD);
            dispatchPage.Filter(DispatchPage.FilterType.Search, delivery4Name);
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
            // dispatch 5 NO DELIVERY ROUND
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);
            dispatchPage.Filter(DispatchPage.FilterType.Search, delivery5Name);
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
        }
        public void PR_SETUP_ValidateDispatch_TestInitialize()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // dispatch 2
            DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryRound2Name);
            PrevisionalQtyPage previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
            // dispatch 3
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteMAD);
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryRound3Name);
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
            // dispatch 4
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteMAD);
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryRound4Name);
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
            // dispatch 5 NO DELIVERY ROUND
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);
            dispatchPage.Filter(DispatchPage.FilterType.Search, delivery5Name);
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
        }

        public void PR_SETUP_Create_Datas_MethodIndividualMultiportion_TestInitialize()
        {
            // Prepare            
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();
            string qty = "10";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();

            // Act
            //Create Customer ES
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerESName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerESName, customerESCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerESName);
            }

            Assert.IsTrue(customersPage.GetFirstCustomerName().Contains(customerESName), "Le customer n'a pas été créé.");

            var servicePage = homePage.GoToCustomers_ServicePage();
            // SERVICIO ES IND
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESINDServicioName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(ESINDServicioName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerESName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerESName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESINDServicioName);

            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(ESINDServicioName), "Le service " + ESINDServicioName + " n'existe pas.");

            // SERVICIO ES MUL
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESMULServicioName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(ESMULServicioName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerESName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerESName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESMULServicioName);

            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(ESMULServicioName), "Le service " + ESMULServicioName + " n'existe pas.");

            // SERVICIO INDMUL
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESINDMULServicioName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(ESINDMULServicioName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerESName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerESName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESINDMULServicioName);

            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(ESINDMULServicioName), "Le service " + ESINDMULServicioName + " n'existe pas.");

            //Create Receta ES
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaESName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recetaESName, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaESName);
            Assert.AreEqual(recetaESName, recipesPage.GetFirstRecipeName(), "La recette " + recetaESName + " n'existe pas.");

            //Create menu ES
            // menu ES IND
            var menusPage = homePage.GoToMenus_Menus();

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESINDMenuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(ESINDServicioName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(ESINDMenuName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaESName);
                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeCoef(ESINDCoefficient);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddDays(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESINDCoefficient);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESINDCoefficient);
                }

                menusPage = menuDayViewPage.BackToList();

            }
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESINDMenuName);
            Assert.AreEqual(ESINDMenuName, menusPage.GetFirstMenuName(), "Le menu1 n'a pas été créé.");

            // menu ES MUL
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESMULMenuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(ESMULServicioName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(ESMULMenuName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaESName);
                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeCoef(ESMULCoefficient);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESMULCoefficient);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESMULCoefficient);
                }
                menusPage = menuDayViewPage.BackToList();

            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESMULMenuName);
            Assert.AreEqual(ESMULMenuName, menusPage.GetFirstMenuName(), "Le menu2 n'a pas été créé.");

            // menu ES IND MUL
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESINDMULMenuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(ESINDMULServicioName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(ESINDMULMenuName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaESName);
                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeCoef(ESINDMULCoefficient);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESINDMULCoefficient);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESINDMULCoefficient);
                }

                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESINDMULMenuName);
            Assert.AreEqual(ESINDMULMenuName, menusPage.GetFirstMenuName(), "Le menu n'a pas été créé.");


            //Create Delivery ES
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, ESDeliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(ESDeliveryName, customerESName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio ES IND
                deliveryLoadingPage.AddService(ESINDServicioName);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethodIndividualMultiporcion, foodpackIndividual, false, ESINDServicioPaxMin, ESINDServicioPaxMax, ESINDServicioPortionsNumber);
                deliveryLoadingPage.CloseFoodPackagingModal();

                // add servicio ES MUL
                deliveryLoadingPage.AddService(ESMULServicioName);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethodIndividualMultiporcion, foodpackMultiporcion, false, ESMULServicioPaxMin, ESMULServicioPaxMax, ESMULServicioPortionsNumber);
                deliveryLoadingPage.CloseFoodPackagingModal();

                // add servicio ES IND MUL
                deliveryLoadingPage.AddService(ESINDMULServicioName);
                deliveryLoadingPage.AddQuantity(qty);
                // add individual packaging
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethodIndividualMultiporcion, foodpackIndividual, false, ESINDMULIndividualServicioPaxMin, ESINDMULIndividualServicioPaxMax, ESINDMULIndividualServicioPortionsNumber);

                // add multiporcion packaging
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethodIndividualMultiporcion, foodpackMultiporcion, false, ESINDMULMultiporcionServicioPaxMin, ESINDMULMultiporcionServicioPaxMax, ESINDMULMultiporcionServicioPortionsNumber);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, ESDeliveryName);
            }

            //Create delivery round ES
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundES);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRoundES, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(ESDeliveryName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundES);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(deliveryRoundES), "Le delivery round n'a pas été créé.");
        }

        public void PR_SETUP_Create_Datas_MethodPaxPerPackaging_TestInitialize()
        {
            // Prepare            
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();

            // Act
            //Create Customer TLS
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, TLSCustomerName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(TLSCustomerName, TLSCustomerCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, TLSCustomerName);
            }

            var customerTLSExactName = customersPage.GetFirstCustomerName();
            Assert.IsTrue(customerTLSExactName.Contains(TLSCustomerName), "Le customer n'a pas été créé.");

            //Create SERVICIO TLS
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, TLSServicioName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(TLSServicioName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, TLSCustomerName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, TLSCustomerName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, TLSServicioName);

            Assert.IsTrue(TLSServicioName.Contains(TLSServicioName), "Le service " + TLSServicioName + " n'existe pas.");

            //Create Receta TLS
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, TLSRecetaName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(TLSRecetaName, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, TLSRecetaName);
            Assert.IsTrue(TLSRecetaName.Contains(TLSRecetaName), "La recette " + TLSRecetaName + " n'existe pas.");

            //Create menu TLS
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, TLSMenuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(TLSServicioName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(TLSMenuName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(TLSRecetaName);
                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeCoef(TLSCoefficient200);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(TLSRecetaName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(TLSCoefficient200);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(TLSRecetaName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(TLSCoefficient200);
                }

                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, TLSMenuName);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(TLSMenuName), "Le menu n'a pas été créé.");

            //Create Delivery TLS
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, TLSDeliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(TLSDeliveryName, TLSCustomerName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio TLS
                deliveryLoadingPage.AddService(TLSServicioName);
                deliveryLoadingPage.AddQuantity("10");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB1, false, null, null, portionsFoodpackB1);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB4, false, null, null, portionsFoodpackB4);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB6, false, null, null, portionsFoodpackB6);
                deliveryLoadingPage.CloseFoodPackagingModal();
                deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, TLSDeliveryName);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(TLSDeliveryName), "La delivery n'a pas été crée");

            //Create delivery round TLS
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, TLSDeliveryRound);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(TLSDeliveryRound, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(TLSDeliveryName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, TLSDeliveryRound);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(TLSDeliveryRound), "Le delivery round n'a pas été créé.");
        }

        public void PR_SETUP_Create_Datas_MethodPaxPerPackagingGrouped_TestInitialize()
        {
            // Prepare            
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();

            // Act
            //Create Customer TLS
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerTLSGRName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerTLSGRName, customerTLSGRCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerTLSGRName);
            }

            //Create SERVICIO TLS GR 1
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioTLSGR1Name);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(servicioTLSGR1Name, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerTLSGRName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerTLSGRName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioTLSGR1Name);

            var servicioExactName = servicePage.GetFirstServiceName();
            Assert.IsTrue(servicioExactName.Contains(servicioTLSGR1Name), "Le service " + servicioTLSGR1Name + " n'existe pas.");

            //Create SERVICIO TLS GR 2
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioTLSGR2Name);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(servicioTLSGR2Name, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerTLSGRName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerTLSGRName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioTLSGR2Name);

            var servicio2xactName = servicePage.GetFirstServiceName();
            Assert.IsTrue(servicio2xactName.Contains(servicioTLSGR2Name), "Le service " + servicioTLSGR2Name + " n'existe pas.");

            //Create Receta TLS
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaTLSGRName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recetaTLSGRName, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion100g);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion100g);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaTLSGRName);
            var recetaExactName = recipesPage.GetFirstRecipeName();
            Assert.IsTrue(recetaExactName.Contains(recetaTLSGRName), "La recette " + recetaTLSGRName + " n'existe pas.");

            //Create menu TLS GR 1
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuTLSGR1Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(servicioTLSGR1Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuTLSGR1Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaTLSGRName);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaTLSGRName);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaTLSGRName);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuTLSGR1Name);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(menuTLSGR1Name), "Le menu n'a pas été créé.");

            //Create menu TLS GR 2
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuTLSGR2Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(servicioTLSGR2Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuTLSGR2Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaTLSGRName);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaTLSGRName);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaTLSGRName);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuTLSGR2Name);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(menuTLSGR2Name), "Le menu n'a pas été créé.");

            //Create Delivery TLS GR
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryTLSGRName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryTLSGRName, customerTLSGRName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio TLS GR 1
                deliveryLoadingPage.AddService(servicioTLSGR1Name);
                deliveryLoadingPage.AddQuantity("10");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB1, false, null, null, portionsFoodpackB1);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB4, false, null, null, portionsFoodpackB4);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB6, false, null, null, portionsFoodpackB6);
                deliveryLoadingPage.CloseFoodPackagingModal();

                // add servicio TLS GR 2 
                deliveryLoadingPage.AddService(servicioTLSGR2Name);
                deliveryLoadingPage.AddQuantity("10");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB1, false, null, null, portionsFoodpackB1);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB4, false, null, null, portionsFoodpackB4);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB6, false, null, null, portionsFoodpackB6);
                deliveryLoadingPage.CloseFoodPackagingModal();

                // change method to PaxPerPackGrouped
                var deliveryGeneralInformationPage = deliveryLoadingPage.ClickOnGeneralInformation();
                deliveryGeneralInformationPage.SetMethod("PaxPerPackGrouped");

                deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryTLSGRName);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(deliveryTLSGRName), "La delivery n'a pas été crée");

            //Create delivery round TLS
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundTLSGR);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRoundTLSGR, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryTLSGRName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundTLSGR);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(deliveryRoundTLSGR), "Le delivery round n'a pas été créé.");

            //Dispatch DELIVERY WITH SERVICIO TLS GR 1 = 5 & DELIVERY WITH SERVICIO TLS GR 2 = 10
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, "SERVICIO GR TLS");
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(servicioTLSGR1Name, TLSGRServicio1DispatchQuantity);
            dispatchPage.AddQuantityOnPrevisonalQuantity(servicioTLSGR2Name, TLSGRServicio2DispatchQuantity);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(servicioTLSGR1Name)) > 0, "Le service {0} n'a pas de menu associé.", servicioTLSGR1Name);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(servicioTLSGR2Name)) > 0, "Le service {0} n'a pas de menu associé.", servicioTLSGR2Name);
        }

        public void PR_SETUP_Create_Datas_MethodByRecipeVariant_TestInitialize()
        {
            // Prepare
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            ParametersSites sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();
            // Act
            //Create Customer PF
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerPFName);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerPFName, customerPFCode, customerType1);
                CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerPFName);
            }
            bool isCreated = customersPage.GetFirstCustomerName().Equals(customerPFName);
            Assert.IsTrue(isCreated, "Le customer n'a pas été créé.");
            //Create SERVICIO PF
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioPFName);
            if (servicePage.CheckTotalNumber() == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(servicioPFName, null, null, serviceCategorie, null, serviceType);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                serviceGeneralInformationsPage.SetProduced(true);
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerPFName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                ServicePricePage pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                ServiceCreatePriceModalPage serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerPFName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioPFName);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(servicioPFName), "Le service " + servicioPFName + " n'existe pas.");

            //Create Receta PF
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaPFName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recetaPFName, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);
                recipeVariantPage.ClickOnFirstIngredient();
                recipeVariantPage.SetYield(yieldPF95);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);
                recipeVariantPage.ClickOnFirstIngredient();
                recipeVariantPage.SetYield(yieldPF95);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaPFName);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(recetaPFName), "La recette " + recetaPFName + " n'existe pas.");

            //Create menu PF
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuPFName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(servicioPFName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuPFName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaPFName);
                menuDayViewPage.ClickOnFirstRecipe();
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaPFName);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaPFName);
                }

                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuPFName);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(menuPFName), "Le menu {0} n'a pas été créé.", menuPFName);

            var parametersSetup = homePage.GoToParameters_SetupPage();
            parametersSetup.GoToTab_PAXPerRecipeVariant();
            parametersSetup.Filter(ParametersSetup.FilterType.SearchRecipe, recetaPFName);
            parametersSetup.SelectAll();
            parametersSetup.ClickOnMultipleUpdate();
            parametersSetup.AddFoodPackPerRecipe(foodpackBACGN, portionsFoodpackBACGN);
            parametersSetup.SelectAll();
            parametersSetup.ClickOnMultipleUpdate();
            parametersSetup.AddFoodPackPerRecipe(foodpackBACGNhalf, portionsFoodpackBACGNhalf);

            //Create Delivery PF
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryPFName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryPFName, customerPFName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio PF
                deliveryLoadingPage.AddService(servicioPFName);
                deliveryLoadingPage.AddQuantity("16");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackBACGN, false, null, null, portionsFoodpackBACGN);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackBACGNhalf, false, null, null, portionsFoodpackBACGNhalf);
                deliveryLoadingPage.CloseFoodPackagingModal();
                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryPFName);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(deliveryPFName), "La delivery {0} n'a pas été crée", deliveryPFName);

            //Create delivery round PF
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundPF);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRoundPF, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryPFName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundPF);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(deliveryRoundPF), "Le delivery round {0} n'a pas été créé.", deliveryRoundPF);

            //DISPATCH DELIVERY WITH SERVICIO PF = 16
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, servicioPFName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(servicioPFName)) > 0, "Le service {0} n'a pas de menu associé.", servicioPFName);
        }

        public void PR_SETUP_Create_Datas_MethodByRecipeVariant2Recipes_TestInitialize()
        {
            // Prepare
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeType2 = TestContext.Properties["Production_RecipeType2"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();

            // Act
            //Create Customer Condren
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, condrenCustomerName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(condrenCustomerName, condrenCustomerCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, condrenCustomerName);
            }

            Assert.IsTrue(customersPage.GetFirstCustomerName().Equals(condrenCustomerName), "Le customer n'a pas été créé.");

            //Create SERVICIO Condren
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, condrenServicioName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(condrenServicioName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, condrenCustomerName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, condrenCustomerName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, condrenServicioName);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(condrenServicioName), "Le service " + condrenServicioName + " n'existe pas.");

            //Create Receta Condren
            //Receta Condren 1 recipe type = PLATO CALIENTE variante Adulto - almuerzo - poids de portion = 150g
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, condrenReceta1Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(condrenReceta1Name, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);
                recipeVariantPage.ClickOnFirstIngredient();
                recipeVariantPage.SetYield(yieldPF95);
                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);
                recipeVariantPage.WaitPageLoading();
                recipeVariantPage.ClickOnFirstIngredient();
                recipeVariantPage.SetYield(yieldPF95);
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, condrenReceta1Name);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(condrenReceta1Name), "La recette " + condrenReceta1Name + " n'existe pas.");

            //Receta Condren 2 recipe type = ENSALADA variante Adulto - almuerzo - poides de portion = 100g
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, condrenReceta2Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(condrenReceta2Name, recipeType2, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion100g);
                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion100g);
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, condrenReceta2Name);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(condrenReceta2Name), "La recette " + condrenReceta2Name + " n'existe pas.");

            //Create menu Condren RECETA CONDREN 1 methode %  coef 100% + RECETA CONDREN 2 methode std coef 2
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, condrenMenuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(condrenServicioName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(condrenMenuName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(condrenReceta2Name);
                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeMethod("Std");
                menuDayViewPage.SetRecipeCoef("2");
                menuDayViewPage.AddRecipe(condrenReceta1Name);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsThisRecipeDisplayed(condrenReceta2Name))
                {
                    menuDayViewPage.AddRecipe(condrenReceta2Name);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeMethod("Std");
                    menuDayViewPage.SetRecipeCoef("2");
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(condrenReceta1Name))
                {
                    menuDayViewPage.AddRecipe(condrenReceta1Name);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsThisRecipeDisplayed(condrenReceta2Name))
                {
                    menuDayViewPage.AddRecipe(condrenReceta2Name);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeMethod("Std");
                    menuDayViewPage.SetRecipeCoef("2");
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(condrenReceta1Name))
                {
                    menuDayViewPage.AddRecipe(condrenReceta1Name);
                }

                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, condrenMenuName);
            menusPage.WaitPageLoading();


            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(condrenMenuName), "Le menu n'a pas été créé.");

            var parametersSetup = homePage.GoToParameters_SetupPage();
            parametersSetup.GoToTab_PAXPerRecipeVariant();
            parametersSetup.Filter(ParametersSetup.FilterType.SearchRecipe, condrenReceta1Name);
            parametersSetup.SelectAll();
            parametersSetup.ClickOnMultipleUpdate();
            parametersSetup.RemoveFoodPackPerRecipe(foodpackBACGN);
            parametersSetup.RemoveFoodPackPerRecipe(foodpackBACGNhalf);
            parametersSetup.AddFoodPackPerRecipe(foodpackBACGN, portionsFoodpackBACGN);
            parametersSetup.SelectAll();
            parametersSetup.ClickOnMultipleUpdate();
            parametersSetup.AddFoodPackPerRecipe(foodpackBACGNhalf, portionsFoodpackBACGNhalf);

            //Create Delivery CONDREN
            //BY RECIPE VARIANT, recipe types = PLATO CALIENTE, packaging name = BAC G / N & packaging name = BAC G / N 1 / 2
            //PAX PER PACKAGING 2 recipe type = ENSALADA  et BOCADILLO/ SANDWICH, packaging name = B1 - nb of portion = 1

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, condrenDeliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(condrenDeliveryName, condrenCustomerName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio CONDREN
                deliveryLoadingPage.AddService(condrenServicioName);
                deliveryLoadingPage.AddQuantity("15");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGN, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGNhalf, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB1, false, null, null, portionsFoodpackB1, recipeType2);
                deliveryLoadingPage.CloseFoodPackagingModal();
                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, condrenDeliveryName);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(condrenDeliveryName), "La delivery n'a pas été crée");

            //Create delivery round CONDREN
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, condrenDeliveryRound);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(condrenDeliveryRound, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(condrenDeliveryName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, condrenDeliveryRound);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(condrenDeliveryRound), "Le delivery round n'a pas été créé.");

            //DISPATCH DELIVERY WITH SERVICIO CONDREN = 15
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, condrenServicioName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(condrenServicioName)) > 0, "Le service {0} n'a pas de menu associé.", condrenServicioName);
        }

        public void PR_SETUP_Create_Datas_MethodRecipeVariantAndGroupedByRecipe_TestInitialize()
        {
            // Prepare            
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeVariantCollegioForACE = TestContext.Properties["Production_RecipeVariant2ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();
            string menuVariantCollegioForACE = TestContext.Properties["Production_MenuVariant2ForACE"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();

            // Act
            //Create Customer TLS
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, barentinCustomerName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(barentinCustomerName, barentinCustomerCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, barentinCustomerName);
            }

            //Create SERVICIO BARENTIN ADULTO
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, barentinServicioAdulto);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(barentinServicioAdulto, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, barentinCustomerName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, barentinCustomerName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, barentinServicioAdulto);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(barentinServicioAdulto), "Le service " + barentinServicioAdulto + " n'existe pas.");

            //Create SERVICIO BARENTIN COLLEGIO
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, barentinServicioCollegio);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(barentinServicioCollegio, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, barentinCustomerName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, barentinCustomerName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, barentinServicioCollegio);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(barentinServicioCollegio), "Le service " + barentinServicioCollegio + " n'existe pas.");

            //Create Receta Barentin 1
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinReceta1);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(barentinReceta1, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantCollegioForACE);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                recipeVariantPage2.AddIngredient(itemTest);
                recipeVariantPage2.SetTotalPortion(portion100g);
                recipesPage = recipeVariantPage2.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();

                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                if (!recipeVariantPage2.IsIngredientDisplayed())
                {
                    recipeVariantPage2.AddIngredient(itemTest);
                }
                recipeVariantPage2.SetTotalPortion(portion100g);

                recipesPage = recipeVariantPage2.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinReceta1);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(barentinReceta1), "La recette " + barentinReceta1 + " n'existe pas.");

            //Create Receta Barentin 2
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinRecetaNP2);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(barentinRecetaNP2, recipeTypeNP, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantCollegioForACE);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                recipeVariantPage2.AddIngredient(itemTest);
                recipeVariantPage2.SetTotalPortion(portion100g);
                recipesPage = recipeVariantPage2.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();

                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                if (!recipeVariantPage2.IsIngredientDisplayed())
                {
                    recipeVariantPage2.AddIngredient(itemTest);
                }
                recipeVariantPage2.SetTotalPortion(portion100g);

                recipesPage = recipeVariantPage2.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinRecetaNP2);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(barentinRecetaNP2), "La recette " + barentinReceta1 + " n'existe pas.");

            //Create Receta Barentin 3
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinRecetaBulk3);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(barentinRecetaBulk3, recipeTypeBulk, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantCollegioForACE);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                recipeVariantPage2.AddIngredient(itemTest);
                recipeVariantPage2.SetTotalPortion(portion100g);
                recipesPage = recipeVariantPage2.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();

                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                if (!recipeVariantPage2.IsIngredientDisplayed())
                {
                    recipeVariantPage2.AddIngredient(itemTest);
                }
                recipeVariantPage2.SetTotalPortion(portion100g);

                recipesPage = recipeVariantPage2.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinRecetaBulk3);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(barentinRecetaBulk3), "La recette " + barentinRecetaBulk3 + " n'existe pas.");

            //Create Menu barentin adulto
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, barentinMenuAdulto);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(barentinServicioAdulto);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(barentinMenuAdulto, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(barentinReceta1);
                menuDayViewPage.AddRecipe(barentinRecetaNP2);
                menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinReceta1))
                {
                    menuDayViewPage.AddRecipe(barentinReceta1);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaNP2))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaNP2);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaBulk3))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinReceta1))
                {
                    menuDayViewPage.AddRecipe(barentinReceta1);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaNP2))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaNP2);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaBulk3))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, barentinMenuAdulto);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(barentinMenuAdulto), "Le menu n'a pas été créé.");

            //Create Menu barentin collegio
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, barentinMenuCollegio);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(barentinServicioCollegio);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(barentinMenuCollegio, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantCollegioForACE);
                menuDayViewPage.AddRecipe(barentinReceta1);
                menuDayViewPage.AddRecipe(barentinRecetaNP2);
                menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinReceta1))
                {
                    menuDayViewPage.AddRecipe(barentinReceta1);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaNP2))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaNP2);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaBulk3))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinReceta1))
                {
                    menuDayViewPage.AddRecipe(barentinReceta1);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaNP2))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaNP2);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaBulk3))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, barentinMenuCollegio);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(barentinMenuCollegio), "Le menu n'a pas été créé.");

            var parametersSetup = homePage.GoToParameters_SetupPage();
            parametersSetup.GoToTab_PAXPerRecipeVariant();

            parametersSetup.AddPaxPerRecipe(barentinReceta1, menuVariantForACE, foodpackBACGN, portionsFoodpackBACGN, false);
            parametersSetup.AddPaxPerRecipe(barentinReceta1, menuVariantForACE, foodpackBACGNhalf, portionsFoodpackBACGNhalf, false);

            parametersSetup.AddPaxPerRecipe(barentinReceta1, menuVariantCollegioForACE, foodpackBACGN, "16", false);
            parametersSetup.AddPaxPerRecipe(barentinReceta1, menuVariantCollegioForACE, foodpackBACGNhalf, "8", false);

            parametersSetup.AddPaxPerRecipe(barentinRecetaBulk3, menuVariantForACE, foodpackEntier, "6", true);
            parametersSetup.AddPaxPerRecipe(barentinRecetaBulk3, menuVariantForACE, foodpackDemi, "3", true);

            parametersSetup.AddPaxPerRecipe(barentinRecetaBulk3, menuVariantCollegioForACE, foodpackEntier, "8", true);
            parametersSetup.AddPaxPerRecipe(barentinRecetaBulk3, menuVariantCollegioForACE, foodpackDemi, "4", true);

            parametersSetup.AddPaxPerRecipe(barentinRecetaNP2, menuVariantForACE, foodpackCRT, "50");
            parametersSetup.AddPaxPerRecipe(barentinRecetaNP2, menuVariantCollegioForACE, foodpackCRT, "50");

            //Settings/Set Up/ ajout PAX Per Recipe sur la recipe Bulk
            parametersSetup.GoToTab_PAXPerRecipe();
            parametersSetup.AddPaxPerRecipe(barentinRecetaNP2, null, foodpackCRT, "50");
            parametersSetup.AddPaxPerRecipe(barentinRecetaNP2, null, foodpackUnit, "1");
            parametersSetup.AddPaxPerRecipe(barentinRecetaNP2, null, foodpackSEAU, "10");

            //Create Delivery barentin
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, barentinDeliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(barentinDeliveryName, barentinCustomerName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio adulto
                deliveryLoadingPage.AddService(barentinServicioAdulto);
                deliveryLoadingPage.AddQuantity("10");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGN, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGNhalf, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackEntier, false, null, null, null, recipeTypeBulk);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackDemi, false, null, null, null, recipeTypeBulk);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackCRT, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackSEAU, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackUnit, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.CloseFoodPackagingModal();

                // add servicio collegio
                deliveryLoadingPage.AddService(barentinServicioCollegio);
                deliveryLoadingPage.AddQuantity("50");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGN, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGNhalf, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackEntier, false, null, null, null, recipeTypeBulk);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackDemi, false, null, null, null, recipeTypeBulk);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackCRT, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackSEAU, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackUnit, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, barentinDeliveryName);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(barentinDeliveryName), "La delivery n'a pas été crée");

            //Create delivery round barentin
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, barentinDeliveryRound);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(barentinDeliveryRound, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(barentinDeliveryName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, barentinDeliveryRound);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(barentinDeliveryRound), "Le delivery round n'a pas été créé.");

            //Dispatch DELIVERY WITH SERVICIO Barentin adulto (10) et collegio (50)
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, "BARENTIN");
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(barentinServicioAdulto)) > 0, "Le service {0} n'a pas de menu associé.", barentinServicioAdulto);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(barentinServicioCollegio)) > 0, "Le service {0} n'a pas de menu associé.", barentinServicioCollegio);
        }



        [TestMethod]
        [Priority(0)]
        [Timeout(Timeout)]
        // TODO: ⚠︎⚠︎⚠︎THIS CODE NEEDS CODE REFACTORING !⚠︎⚠︎⚠︎
        public void PR_SETUP_SetProductionSettings()
        {
            //Variants
            string guestTypeVariant1 = TestContext.Properties["Production_GuestType1"].ToString();
            string guestTypeVariant2 = TestContext.Properties["Production_GuestType2"].ToString();
            string mealVariant1 = TestContext.Properties["Production_Meal1"].ToString();
            string mealVariant2 = TestContext.Properties["Production_Meal2"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            ClearCache();
            ParametersProduction parametersProduction = homePage.GoToParameters_ProductionPage();
            parametersProduction.GoToTab_Variant();
            // Variant 1
            parametersProduction.FilterSite(siteACE);
            parametersProduction.FilterGuestType(guestTypeVariant1);
            parametersProduction.FilterMealType(mealVariant1);

            if (!parametersProduction.IsVariantPresent(guestTypeVariant1, mealVariant1))
            {
                parametersProduction.AddNewVariant(guestTypeVariant1, mealVariant1, siteACE);
            }

            parametersProduction.FilterSite(siteACE);
            parametersProduction.FilterGuestType(guestTypeVariant1);
            parametersProduction.FilterMealType(mealVariant1);

            Assert.IsTrue(parametersProduction.IsVariantPresent(guestTypeVariant1, mealVariant1));

            // Variant 2
            parametersProduction.FilterSite(siteMAD);
            parametersProduction.FilterGuestType(guestTypeVariant2);
            parametersProduction.FilterMealType(mealVariant2);

            if (!parametersProduction.IsVariantPresent(guestTypeVariant2, mealVariant2))
            {
                parametersProduction.AddNewVariant(guestTypeVariant2, mealVariant2, siteMAD);
            }

            parametersProduction.FilterSite(siteMAD);
            parametersProduction.FilterGuestType(guestTypeVariant2);
            parametersProduction.FilterMealType(mealVariant2);

            Assert.IsTrue(parametersProduction.IsVariantPresent(guestTypeVariant2, mealVariant2));

            //Create FoodPacks
            parametersProduction.GoToTab_Foodpack();

            if (!parametersProduction.CheckFoodPackExists(foodpackIndividual))
            {
                parametersProduction.CreateNewFoodPack(foodpackIndividual);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackMultiporcion))
            {
                parametersProduction.CreateNewFoodPack(foodpackMultiporcion);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackB1))
            {
                parametersProduction.CreateNewFoodPack(foodpackB1);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackB4))
            {
                parametersProduction.CreateNewFoodPack(foodpackB4);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackB6))
            {
                parametersProduction.CreateNewFoodPack(foodpackB6);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackBACGN))
            {
                parametersProduction.CreateNewFoodPack(foodpackBACGN);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackBACGNhalf))
            {
                parametersProduction.CreateNewFoodPack(foodpackBACGNhalf);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackDemi))
            {
                parametersProduction.CreateNewFoodPack(foodpackDemi);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackEntier))
            {
                //Add Bulk Equivalence to foodpack
                parametersProduction.CreateNewFoodPack(foodpackEntier);
                parametersProduction.AddBulkEquivalenceToFoodPack(foodpackEntier, foodpackDemi, "2");
                var bulkEquivalence = parametersProduction.GetBulkEquivalence(foodpackEntier);
                Assert.IsTrue(new[] { foodpackDemi, "3" }.Any(eq => bulkEquivalence.Contains(eq)), "L'équivalence pour le foodpack {0} n'est pas correcte : {1} au lieu de {2} x {3}.", foodpackEntier, bulkEquivalence, "2", foodpackDemi);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackUnit))
            {
                parametersProduction.CreateNewFoodPack(foodpackUnit);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackCRT))
            {
                parametersProduction.CreateNewFoodPack(foodpackCRT);
            }

            if (!parametersProduction.CheckFoodPackExists(foodpackSEAU))
            {
                parametersProduction.CreateNewFoodPack(foodpackSEAU);
            }

            //Create Recipe Type
            parametersProduction.GoToTab_RecipeType();
            if (!parametersProduction.CheckRecipeTypeExists(recipeTypeNP))
            {
                parametersProduction.CreateNewRecipeType(recipeTypeNP, "16");
            }

            if (!parametersProduction.CheckRecipeTypeExists(recipeTypeBulk))
            {
                parametersProduction.CreateNewRecipeType(recipeTypeBulk, "17");
            }
        }

        //Same priority tests than PR_PR (just check in case PR_PR priority tests don't pass)
        [TestMethod]
        [Priority(1)]
        [Timeout(Timeout)]
        // TODO: ⚠︎⚠︎⚠︎THIS CODE NEEDS CODE REFACTORING !⚠︎⚠︎⚠︎ 
        public void PR_SETUP_Create_Customers()
        {
            //warn : dependances deliveries filtres
            // Prepare            
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string customerType2 = TestContext.Properties["Production_CustomerType2"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            //second customer for Production
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer2Name);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer2Name, customer2Code, customerType2);
                CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer2Name);
            }
            string firstCustomerName = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customer2Name, firstCustomerName, "Le customer n'a pas été créé.");
            //third customer for Production
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer3Name);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer3Name, customer3Code, customerType1);
                CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer3Name);
            }
            firstCustomerName = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customer3Name, firstCustomerName, "Le customer n'a pas été créé.");
            //fourth customer for Production
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer4Name);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer4Name, customer4Code, customerType2);
                CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer4Name);
            }
            firstCustomerName = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customer4Name, firstCustomerName, "Le customer n'a pas été créé.");
            //fifth customer for Production (no delivery round)
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer5Name);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer5Name, customer5Code, customerType2);
                CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer5Name);
            }
            firstCustomerName = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customer5Name, firstCustomerName, "Le customer n'a pas été créé.");
        }

        [TestMethod]
        [Priority(2)]
        [Timeout(Timeout)]
        public void PR_SETUP_Create_Services()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string customer2Reference = $"{customer2Code} - {customer2Name}";
            string customer3Reference = $"{customer3Code} - {customer3Name}";
            string customer4Reference = $"{customer4Code} - {customer4Name}";
            string customer5Reference = $"{customer5Code} - {customer5Name}";
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            // service 2
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service2Name);
            if (servicePage.CheckTotalNumber() == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(service2Name, null, null, serviceCategorie, null, serviceType);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                serviceGeneralInformationsPage.SetProduced(true);
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customer2Reference, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                ServicePricePage pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                ServiceCreatePriceModalPage serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customer2Name);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }
                ServiceGeneralInformationPage serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service2Name);
            string service2 = servicePage.GetFirstServiceName();
            Assert.IsTrue(service2.Contains(service2Name), $"Le service {service2Name} n'existe pas.");

            // service 3
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service3Name);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(service3Name, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteMAD, customer3Reference, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteMAD, customer3Name);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service3Name);

            string service3 = servicePage.GetFirstServiceName();

            Assert.IsTrue(service3.Contains(service3Name), "Le service " + service3Name + " n'existe pas.");

            // service 4
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service4Name);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(service4Name, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteMAD, customer4Reference, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteMAD, customer4Name);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service4Name);

            string service4 = servicePage.GetFirstServiceName();

            Assert.IsTrue(service4.Contains(service4Name), "Le service " + service4Name + " n'existe pas.");

            // service 5 no delivery round
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service5Name);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(service5Name, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customer5Reference, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customer5Name);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service5Name);

            string service5 = servicePage.GetFirstServiceName();

            Assert.IsTrue(service5.Contains(service5Name), "Le service " + service5Name + " n'existe pas.");

        }

        [TestMethod]
        [Priority(3)]
        [Timeout(Timeout)]
        public void PR_SETUP_Create_Items()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            string group1 = TestContext.Properties["Production_ItemGroup1"].ToString();
            string group2 = TestContext.Properties["Production_ItemGroup2"].ToString();

            string workshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string workshop2 = TestContext.Properties["Production_Workshop2"].ToString();

            string supplier1 = TestContext.Properties["Production_ItemSupplier1"].ToString();
            string supplier2 = TestContext.Properties["Production_ItemSupplier2"].ToString();

            string taxType = TestContext.Properties["Production_ItemTaxType"].ToString();
            string prodUnit = TestContext.Properties["Production_ItemProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();

            // ----------------------------------------- Item 2 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, item2Name);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item2Name, group2, workshop2, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier2);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, item2Name);
            }

            string item2 = itemPage.GetFirstItemName();

            // ----------------------------------------- Item 3 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, item3Name);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item3Name, group1, workshop1, taxType, prodUnit);

                // 1 packaging pour le site MAD
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteMAD, packagingName, storageQty, storageUnit, qty, supplier1);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, item3Name);
            }

            string item3 = itemPage.GetFirstItemName();

            // ----------------------------------------- Item 2 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, item4Name);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item4Name, group2, workshop2, taxType, prodUnit);

                // 1 packaging pour le site MAD
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteMAD, packagingName, storageQty, storageUnit, qty, supplier2);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, item4Name);
            }

            string item4 = itemPage.GetFirstItemName();

            // ----------------------------------------- Item 3 ACE NO DELIVERY ROUND ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, item5Name);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item5Name, group2, workshop2, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier2);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, item5Name);
            }

            string item5 = itemPage.GetFirstItemName();

            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemTest);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemTest, group1, workshop1, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier1);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemTest);
            }

            string item6 = itemPage.GetFirstItemName();

            // Assert
            Assert.AreEqual(item2Name, item2, "L'item " + item2Name + " n'existe pas.");
            Assert.AreEqual(item3Name, item3, "L'item " + item3Name + " n'existe pas.");
            Assert.AreEqual(item4Name, item4, "L'item " + item4Name + " n'existe pas.");
            Assert.AreEqual(item5Name, item5, "L'item " + item5Name + " n'existe pas.");
            Assert.AreEqual(itemTest, item6, "L'item " + itemTest + " n'existe pas.");
        }

        [TestMethod]
        [Priority(4)]
        [Timeout(Timeout)]
        public void PR_SETUP_AddItemsToPriceLists()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            PageObjects.Customers.PriceList.PriceListPage priceListPage = homePage.GoToCustomers_PriceListPage();
            // price list 1 for ACE
            priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Site, siteACE);
            priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Search, priceList1Name);
            if (priceListPage.CheckTotalNumber() == 0)
            {
                //création de la price list
                PageObjects.Customers.PriceList.PricingCreateModalpage pricingCreateModalpage = priceListPage.CreateNewPricing();
                PageObjects.Customers.PriceList.PriceListDetailsPage priceListDetailsPage = pricingCreateModalpage.FillField_CreateNewPricing(priceList1Name, siteACE, DateUtils.Now, DateUtils.Now, true);
                priceListDetailsPage.AddItem(item2Name, false);
                priceListDetailsPage.AddItem(item5Name, false);
                priceListDetailsPage.BackToList();
                priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Search, priceList1Name);
            }
            int totalNumber = priceListPage.CheckTotalNumber();
            Assert.IsTrue(totalNumber == 1, $"La priceList : {priceList1Name} n'a pas été crée");
            if (totalNumber == 1)
            {
                //La priceList est crée, on vérifié la présence des items
                PageObjects.Customers.PriceList.PriceListDetailsPage priceListDetailsPage = priceListPage.ClickOnFirstPricing();
                if (!priceListDetailsPage.IsItemAdded(item2Name)) priceListDetailsPage.AddItem(item2Name, false);
                if (!priceListDetailsPage.IsItemAdded(item5Name)) priceListDetailsPage.AddItem(item5Name, false);
                //Asserts
                bool isItemAdded = priceListDetailsPage.IsItemAdded(item2Name);
                Assert.IsTrue(isItemAdded, $"L'item : {item2Name} n'a pas été ajouté à la PriceList : {priceList1Name}");
                isItemAdded = priceListDetailsPage.IsItemAdded(item5Name);
                Assert.IsTrue(isItemAdded, $"L'item : {item5Name} n'a pas été ajouté à la PriceList : {priceList1Name}");
                priceListDetailsPage.BackToList();
            }
            priceListPage.ResetFilter();
            // price list for MAD
            priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Site, siteMAD);
            priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Search, priceList2Name);
            totalNumber = priceListPage.CheckTotalNumber();
            if (totalNumber == 0)
            {
                //création de la price list
                PageObjects.Customers.PriceList.PricingCreateModalpage pricingCreateModalpage = priceListPage.CreateNewPricing();
                PageObjects.Customers.PriceList.PriceListDetailsPage priceListDetailsPage = pricingCreateModalpage.FillField_CreateNewPricing(priceList2Name, siteMAD, DateUtils.Now, DateUtils.Now, true);
                priceListDetailsPage.AddItem(item3Name, false);
                priceListDetailsPage.AddItem(item4Name, false);
                priceListDetailsPage.BackToList();
                priceListPage.Filter(PageObjects.Customers.PriceList.PriceListPage.FilterType.Search, priceList2Name);
            }
            Assert.IsTrue(totalNumber == 1, $"La priceList : {priceList2Name} n'a pas été crée");
            if (totalNumber == 1)
            {
                //La priceList est crée, on vérifié la présence des items
                PageObjects.Customers.PriceList.PriceListDetailsPage priceListDetailsPage = priceListPage.ClickOnFirstPricing();
                if (!priceListDetailsPage.IsItemAdded(item3Name)) priceListDetailsPage.AddItem(item3Name, false);
                if (!priceListDetailsPage.IsItemAdded(item4Name)) priceListDetailsPage.AddItem(item4Name, false);
                //Asserts
                bool isItemAdded = priceListDetailsPage.IsItemAdded(item3Name);
                Assert.IsTrue(isItemAdded, $"L'item : {item3Name} n'a pas été ajouté à la PriceList : {priceList2Name}");
                isItemAdded = priceListDetailsPage.IsItemAdded(item4Name);
                Assert.IsTrue(isItemAdded, $"L'item : {item4Name} n'a pas été ajouté à la PriceList : {priceList2Name}");
            }
        }

        [Priority(5)]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Create_Recipes()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeVariantForMAD = TestContext.Properties["Production_RecipeVariant1ForMAD"].ToString();

            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";

            string recipeType2 = TestContext.Properties["Production_RecipeType2"].ToString();
            string recipeCookingMode2 = TestContext.Properties["Production_CookingMode2"].ToString();
            string recipeWorkshop2 = TestContext.Properties["Production_Workshop2"].ToString();
            string recipePortion2 = "18";

            string newTotalPortion = "100";

            //Arrange           
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();

            // ---------------------------------------------- recette 2 --------------------------------------------------

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe2Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipe2Name, recipeType2, recipePortion2, true, null, recipeCookingMode2, recipeWorkshop2);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(item2Name);
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(item2Name);
                }
                recipeVariantPage.SetTotalPortion(newTotalPortion);
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe2Name);

            string recipe2 = recipesPage.GetFirstRecipeName();

            // ---------------------------------------------- recette 3 --------------------------------------------------

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe3Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipe3Name, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteMAD, recipeVariantForMAD);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(item3Name);
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(item3Name);
                }
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe3Name);

            string recipe3 = recipesPage.GetFirstRecipeName();

            // ---------------------------------------------- recette 4 --------------------------------------------------

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe4Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipe4Name, recipeType2, recipePortion2, true, null, recipeCookingMode2, recipeWorkshop2);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteMAD, recipeVariantForMAD);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(item4Name);

                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(item4Name);
                }
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe4Name);

            string recipe4 = recipesPage.GetFirstRecipeName();

            // ---------------------------------------------- recette 5 NO DELIVERY ROUND --------------------------------------------------

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe5Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipe5Name, recipeType2, recipePortion2, true, null, recipeCookingMode2, recipeWorkshop2);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(item5Name);
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(item5Name);
                }
                recipeVariantPage.SetTotalPortion(newTotalPortion);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe5Name);

            string recipe5 = recipesPage.GetFirstRecipeName();

            // --------------------------------------------- Assert ------------------------------------------------------------
            Assert.AreEqual(recipe2Name, recipe2, "La recette " + recipe2Name + " n'existe pas.");
            Assert.AreEqual(recipe3Name, recipe3, "La recette " + recipe3Name + " n'existe pas.");
            Assert.AreEqual(recipe4Name, recipe4, "La recette " + recipe4Name + " n'existe pas.");
            Assert.AreEqual(recipe5Name, recipe5, "La recette " + recipe5Name + " n'existe pas.");
        }

        [TestMethod]
        [Priority(6)]
        [Timeout(Timeout)]
        public void PR_SETUP_Create_Menus()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            string recipeVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();
            string recipeVariantForMAD = TestContext.Properties["Production_MenuVariant1ForMAD"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();

            // menu 2
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu2Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(service2Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menu2Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, recipeVariantForACE);
                menuDayViewPage.AddRecipe(recipe2Name);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe2Name);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe2Name);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu2Name);
            Assert.AreEqual(menu2Name, menusPage.GetFirstMenuName(), "Le menu2 n'a pas été créé.");


            // menu 3
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu3Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteMAD);
                menusCreateModalPage.AddService(service3Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menu3Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteMAD, recipeVariantForMAD);
                menuDayViewPage.AddRecipe(recipe3Name);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe3Name);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe3Name);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu3Name);
            Assert.AreEqual(menu3Name, menusPage.GetFirstMenuName(), "Le menu n'a pas été créé.");

            // menu 4
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu4Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteMAD);
                menusCreateModalPage.AddService(service4Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menu4Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteMAD, recipeVariantForMAD);
                menuDayViewPage.AddRecipe(recipe4Name);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe4Name);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe4Name);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu4Name);
            Assert.AreEqual(menu4Name, menusPage.GetFirstMenuName(), "Le menu n'a pas été créé.");

            // menu 5 NO DELIVERY ROUND
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu5Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(service5Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menu5Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, recipeVariantForACE);
                menuDayViewPage.AddRecipe(recipe5Name);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe5Name);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipe5Name);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menu5Name);
            Assert.AreEqual(menu5Name, menusPage.GetFirstMenuName(), "Le menu n'a pas été créé.");
        }

        [TestMethod]
        [Priority(7)]
        [Timeout(Timeout)]
        public void PR_SETUP_Create_Deliveries()
        {
            // Prepare
            string qty = "10";
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            string packagingMethod1 = "By recipe-variant";
            string packagingName1 = "Individual";
            string packagingMethod2 = "PAX per packaging";
            string packagingName2 = "Multiporcion";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            //delivery 2
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery2Name);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(delivery2Name, customer2Name, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                deliveryLoadingPage.AddService(service2Name);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethod2, packagingName2, true);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery2Name);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Equals(delivery2Name), "La delivery : " + delivery2Name + " n'a pas été crée");

            deliveryPage.ResetFilter();

            //delivery 3
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery3Name);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(delivery3Name, customer3Name, siteMAD, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                deliveryLoadingPage.AddService(service3Name);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethod1, packagingName2, false);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery3Name);

            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Equals(delivery3Name), "La delivery : " + delivery3Name + " n'a pas été crée");

            deliveryPage.ResetFilter();

            //delivery 4
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery4Name);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(delivery4Name, customer4Name, siteMAD, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                deliveryLoadingPage.AddService(service4Name);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethod2, packagingName1, true);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery4Name);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Equals(delivery4Name), "La delivery : " + delivery4Name + " n'a pas été crée");

            //delivery 5 NO DELIVERY ROUND
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery5Name);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(delivery5Name, customer5Name, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                deliveryLoadingPage.AddService(service5Name);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethod2, packagingName2, true);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, delivery5Name);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Equals(delivery5Name), "La delivery : " + delivery5Name + " n'a pas été crée");
        }

        //Same priority tests than PR_PR (just check in case PR_PR priority tests don't pass)
        [TestMethod]
        [Priority(8)]
        [Timeout(Timeout)]
        public void PR_SETUP_Create_Delivery_Rounds()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();
            ParametersSites sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteMAD);
            string siteIdMAD = sitePage.CollectNewSiteID();

            // Act

            // delivery round 2
            DeliveryRoundPage deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound2Name);
            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                DeliveryRoundCreateModalpage deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRound2Name, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));
                DeliveryRoundDeliveryPage deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(delivery2Name);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound2Name);
            }
            else
            {
                DeliveryRoundDeliveryPage deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }
            //Assert
            string firstDeliveryRound = deliveryRoundPage.GetFirstDeliveryRound();
            Assert.AreEqual(deliveryRound2Name, firstDeliveryRound, "Le delivery round n'a pas été créé.");
            // delivery round 3
            deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound3Name);
            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                DeliveryRoundCreateModalpage deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRound3Name, siteIdMAD, DateUtils.Now, DateUtils.Now.AddDays(+31));
                DeliveryRoundDeliveryPage deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(delivery3Name);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound3Name);
            }
            else
            {
                DeliveryRoundDeliveryPage deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }
            //Assert
            firstDeliveryRound = deliveryRoundPage.GetFirstDeliveryRound();
            Assert.AreEqual(deliveryRound3Name, firstDeliveryRound, "Le delivery round n'a pas été créé.");
            // delivery round 4
            deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound4Name);
            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                DeliveryRoundCreateModalpage deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRound4Name, siteIdMAD, DateUtils.Now, DateUtils.Now.AddDays(+31));
                DeliveryRoundDeliveryPage deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(delivery4Name);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound4Name);
            }
            else
            {
                DeliveryRoundDeliveryPage deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }
            firstDeliveryRound = deliveryRoundPage.GetFirstDeliveryRound();
            //Assert
            Assert.AreEqual(deliveryRound4Name, firstDeliveryRound, "Le delivery round n'a pas été créé.");
        }

        [TestMethod]
        [Priority(9)]
        [Timeout(Timeout)]
        // TODO: ⚠︎⚠︎⚠︎THIS CODE NEEDS CODE REFACTORING !⚠︎⚠︎⚠︎ 
        public void PR_SETUP_ValidateDispatchs()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // dispatch 2
            DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);
            dispatchPage.Filter(DispatchPage.FilterType.Search, delivery2Name);
            PrevisionalQtyPage previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
            // dispatch 3
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteMAD);
            dispatchPage.Filter(DispatchPage.FilterType.Search, delivery3Name);
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
            // dispatch 4
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteMAD);
            dispatchPage.Filter(DispatchPage.FilterType.Search, delivery4Name);
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
            // dispatch 5 NO DELIVERY ROUND
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);
            dispatchPage.Filter(DispatchPage.FilterType.Search, delivery5Name);
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            isValidate = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isValidate, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
        }

        //Create Datas Delivery ES (packaging individual/multiporcion, 3 services)
        [TestMethod]
        [Priority(10)]
        [Timeout(Timeout)]
        public void PR_SETUP_Create_Datas_MethodIndividualMultiportion()
        {
            // Prepare            
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();
            string qty = "10";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();

            // Act
            //Create Customer ES
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerESName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerESName, customerESCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerESName);
            }

            Assert.IsTrue(customersPage.GetFirstCustomerName().Contains(customerESName), "Le customer n'a pas été créé.");

            var servicePage = homePage.GoToCustomers_ServicePage();
            // SERVICIO ES IND
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESINDServicioName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(ESINDServicioName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerESName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerESName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESINDServicioName);

            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(ESINDServicioName), "Le service " + ESINDServicioName + " n'existe pas.");

            // SERVICIO ES MUL
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESMULServicioName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(ESMULServicioName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerESName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerESName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESMULServicioName);

            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(ESMULServicioName), "Le service " + ESMULServicioName + " n'existe pas.");

            // SERVICIO INDMUL
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESINDMULServicioName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(ESINDMULServicioName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerESName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerESName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ESINDMULServicioName);

            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(ESINDMULServicioName), "Le service " + ESINDMULServicioName + " n'existe pas.");

            //Create Receta ES
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaESName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recetaESName, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaESName);
            Assert.AreEqual(recetaESName, recipesPage.GetFirstRecipeName(), "La recette " + recetaESName + " n'existe pas.");

            //Create menu ES
            // menu ES IND
            var menusPage = homePage.GoToMenus_Menus();

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESINDMenuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(ESINDServicioName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(ESINDMenuName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaESName);
                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeCoef(ESINDCoefficient);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddDays(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESINDCoefficient);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESINDCoefficient);
                }

                menusPage = menuDayViewPage.BackToList();

            }
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESINDMenuName);
            Assert.AreEqual(ESINDMenuName, menusPage.GetFirstMenuName(), "Le menu1 n'a pas été créé.");

            // menu ES MUL
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESMULMenuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(ESMULServicioName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(ESMULMenuName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaESName);
                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeCoef(ESMULCoefficient);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESMULCoefficient);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESMULCoefficient);
                }
                menusPage = menuDayViewPage.BackToList();

            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESMULMenuName);
            Assert.AreEqual(ESMULMenuName, menusPage.GetFirstMenuName(), "Le menu2 n'a pas été créé.");

            // menu ES IND MUL
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESINDMULMenuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(ESINDMULServicioName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(ESINDMULMenuName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaESName);
                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeCoef(ESINDMULCoefficient);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESINDMULCoefficient);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(ESINDMULCoefficient);
                }

                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, ESINDMULMenuName);
            Assert.AreEqual(ESINDMULMenuName, menusPage.GetFirstMenuName(), "Le menu n'a pas été créé.");


            //Create Delivery ES
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, ESDeliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(ESDeliveryName, customerESName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio ES IND
                deliveryLoadingPage.AddService(ESINDServicioName);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethodIndividualMultiporcion, foodpackIndividual, false, ESINDServicioPaxMin, ESINDServicioPaxMax, ESINDServicioPortionsNumber);
                deliveryLoadingPage.CloseFoodPackagingModal();

                // add servicio ES MUL
                deliveryLoadingPage.AddService(ESMULServicioName);
                deliveryLoadingPage.AddQuantity(qty);
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethodIndividualMultiporcion, foodpackMultiporcion, false, ESMULServicioPaxMin, ESMULServicioPaxMax, ESMULServicioPortionsNumber);
                deliveryLoadingPage.CloseFoodPackagingModal();

                // add servicio ES IND MUL
                deliveryLoadingPage.AddService(ESINDMULServicioName);
                deliveryLoadingPage.AddQuantity(qty);
                // add individual packaging
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethodIndividualMultiporcion, foodpackIndividual, false, ESINDMULIndividualServicioPaxMin, ESINDMULIndividualServicioPaxMax, ESINDMULIndividualServicioPortionsNumber);

                // add multiporcion packaging
                deliveryLoadingPage.FillField_FoodPackaging(packagingMethodIndividualMultiporcion, foodpackMultiporcion, false, ESINDMULMultiporcionServicioPaxMin, ESINDMULMultiporcionServicioPaxMax, ESINDMULMultiporcionServicioPortionsNumber);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, ESDeliveryName);
            }

            //Create delivery round ES
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundES);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRoundES, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(ESDeliveryName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundES);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(deliveryRoundES), "Le delivery round n'a pas été créé.");
        }

        //Create Datas Delivery TLS (packaging pax per packaging, 3 foodpacks)
        [TestMethod]
        [Priority(11)]
        [Timeout(Timeout)]
        public void PR_SETUP_Create_Datas_MethodPaxPerPackaging()
        {
            // Prepare            
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();

            // Act
            //Create Customer TLS
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, TLSCustomerName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(TLSCustomerName, TLSCustomerCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, TLSCustomerName);
            }

            var customerTLSExactName = customersPage.GetFirstCustomerName();
            Assert.IsTrue(customerTLSExactName.Contains(TLSCustomerName), "Le customer n'a pas été créé.");

            //Create SERVICIO TLS
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, TLSServicioName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(TLSServicioName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, TLSCustomerName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, TLSCustomerName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, TLSServicioName);

            Assert.IsTrue(TLSServicioName.Contains(TLSServicioName), "Le service " + TLSServicioName + " n'existe pas.");

            //Create Receta TLS
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, TLSRecetaName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(TLSRecetaName, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, TLSRecetaName);
            Assert.IsTrue(TLSRecetaName.Contains(TLSRecetaName), "La recette " + TLSRecetaName + " n'existe pas.");

            //Create menu TLS
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, TLSMenuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(TLSServicioName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(TLSMenuName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(TLSRecetaName);
                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeCoef(TLSCoefficient200);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(TLSRecetaName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(TLSCoefficient200);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(TLSRecetaName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(TLSCoefficient200);
                }

                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, TLSMenuName);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(TLSMenuName), "Le menu n'a pas été créé.");

            //Create Delivery TLS
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, TLSDeliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(TLSDeliveryName, TLSCustomerName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio TLS
                deliveryLoadingPage.AddService(TLSServicioName);
                deliveryLoadingPage.AddQuantity("10");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB1, false, null, null, portionsFoodpackB1);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB4, false, null, null, portionsFoodpackB4);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB6, false, null, null, portionsFoodpackB6);
                deliveryLoadingPage.CloseFoodPackagingModal();
                deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, TLSDeliveryName);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(TLSDeliveryName), "La delivery n'a pas été crée");

            //Create delivery round TLS
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, TLSDeliveryRound);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(TLSDeliveryRound, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(TLSDeliveryName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, TLSDeliveryRound);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(TLSDeliveryRound), "Le delivery round n'a pas été créé.");
        }

        //Create Datas Delivery TLS (packaging pax per packaging, method PaxPerPackGrouped, 3 foodpacks)
        [TestMethod]
        [Priority(12)]
        [Timeout(Timeout)]
        public void PR_SETUP_Create_Datas_MethodPaxPerPackagingGrouped()
        {
            // Prepare            
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();

            // Act
            //Create Customer TLS
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerTLSGRName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerTLSGRName, customerTLSGRCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerTLSGRName);
            }

            //Create SERVICIO TLS GR 1
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioTLSGR1Name);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(servicioTLSGR1Name, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerTLSGRName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerTLSGRName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioTLSGR1Name);

            var servicioExactName = servicePage.GetFirstServiceName();
            Assert.IsTrue(servicioExactName.Contains(servicioTLSGR1Name), "Le service " + servicioTLSGR1Name + " n'existe pas.");

            //Create SERVICIO TLS GR 2
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioTLSGR2Name);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(servicioTLSGR2Name, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerTLSGRName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerTLSGRName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioTLSGR2Name);

            var servicio2xactName = servicePage.GetFirstServiceName();
            Assert.IsTrue(servicio2xactName.Contains(servicioTLSGR2Name), "Le service " + servicioTLSGR2Name + " n'existe pas.");

            //Create Receta TLS
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaTLSGRName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recetaTLSGRName, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion100g);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion100g);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaTLSGRName);
            var recetaExactName = recipesPage.GetFirstRecipeName();
            Assert.IsTrue(recetaExactName.Contains(recetaTLSGRName), "La recette " + recetaTLSGRName + " n'existe pas.");

            //Create menu TLS GR 1
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuTLSGR1Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(servicioTLSGR1Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuTLSGR1Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaTLSGRName);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaTLSGRName);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaTLSGRName);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuTLSGR1Name);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(menuTLSGR1Name), "Le menu n'a pas été créé.");

            //Create menu TLS GR 2
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuTLSGR2Name);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(servicioTLSGR2Name);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuTLSGR2Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaTLSGRName);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaTLSGRName);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaTLSGRName);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuTLSGR2Name);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(menuTLSGR2Name), "Le menu n'a pas été créé.");

            //Create Delivery TLS GR
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryTLSGRName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryTLSGRName, customerTLSGRName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio TLS GR 1
                deliveryLoadingPage.AddService(servicioTLSGR1Name);
                deliveryLoadingPage.AddQuantity("10");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB1, false, null, null, portionsFoodpackB1);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB4, false, null, null, portionsFoodpackB4);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB6, false, null, null, portionsFoodpackB6);
                deliveryLoadingPage.CloseFoodPackagingModal();

                // add servicio TLS GR 2 
                deliveryLoadingPage.AddService(servicioTLSGR2Name);
                deliveryLoadingPage.AddQuantity("10");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB1, false, null, null, portionsFoodpackB1);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB4, false, null, null, portionsFoodpackB4);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB6, false, null, null, portionsFoodpackB6);
                deliveryLoadingPage.CloseFoodPackagingModal();

                // change method to PaxPerPackGrouped
                var deliveryGeneralInformationPage = deliveryLoadingPage.ClickOnGeneralInformation();
                deliveryGeneralInformationPage.SetMethod("PaxPerPackGrouped");

                deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryTLSGRName);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(deliveryTLSGRName), "La delivery n'a pas été crée");

            //Create delivery round TLS
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundTLSGR);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRoundTLSGR, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryTLSGRName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundTLSGR);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(deliveryRoundTLSGR), "Le delivery round n'a pas été créé.");

            //Dispatch DELIVERY WITH SERVICIO TLS GR 1 = 5 & DELIVERY WITH SERVICIO TLS GR 2 = 10
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, "SERVICIO GR TLS");
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(servicioTLSGR1Name, TLSGRServicio1DispatchQuantity);
            dispatchPage.AddQuantityOnPrevisonalQuantity(servicioTLSGR2Name, TLSGRServicio2DispatchQuantity);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(servicioTLSGR1Name)) > 0, "Le service {0} n'a pas de menu associé.", servicioTLSGR1Name);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(servicioTLSGR2Name)) > 0, "Le service {0} n'a pas de menu associé.", servicioTLSGR2Name);

        }

        //Create Datas Delivery PF (packaging by recipe variant, 2 foodpacks)
        [TestMethod]
        [Priority(13)]
        [Timeout(Timeout)]
        public void PR_SETUP_Create_Datas_MethodByRecipeVariant()
        {
            // Prepare
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            ParametersSites sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();
            // Act
            //Create Customer PF
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerPFName);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerPFName, customerPFCode, customerType1);
                CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerPFName);
            }
            bool isCreated = customersPage.GetFirstCustomerName().Equals(customerPFName);
            Assert.IsTrue(isCreated, "Le customer n'a pas été créé.");
            //Create SERVICIO PF
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioPFName);
            if (servicePage.CheckTotalNumber() == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(servicioPFName, null, null, serviceCategorie, null, serviceType);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                serviceGeneralInformationsPage.SetProduced(true);
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerPFName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                ServicePricePage pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                ServiceCreatePriceModalPage serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerPFName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicioPFName);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(servicioPFName), "Le service " + servicioPFName + " n'existe pas.");

            //Create Receta PF
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaPFName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recetaPFName, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);
                recipeVariantPage.ClickOnFirstIngredient();
                recipeVariantPage.SetYield(yieldPF95);

                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);
                recipeVariantPage.ClickOnFirstIngredient();
                recipeVariantPage.SetYield(yieldPF95);

                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaPFName);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(recetaPFName), "La recette " + recetaPFName + " n'existe pas.");

            //Create menu PF
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuPFName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(servicioPFName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuPFName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recetaPFName);
                menuDayViewPage.ClickOnFirstRecipe();
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaPFName);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaPFName);
                }

                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuPFName);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(menuPFName), "Le menu {0} n'a pas été créé.", menuPFName);

            var parametersSetup = homePage.GoToParameters_SetupPage();
            parametersSetup.GoToTab_PAXPerRecipeVariant();
            parametersSetup.Filter(ParametersSetup.FilterType.SearchRecipe, recetaPFName);
            parametersSetup.SelectAll();
            parametersSetup.ClickOnMultipleUpdate();
            parametersSetup.AddFoodPackPerRecipe(foodpackBACGN, portionsFoodpackBACGN);
            parametersSetup.SelectAll();
            parametersSetup.ClickOnMultipleUpdate();
            parametersSetup.AddFoodPackPerRecipe(foodpackBACGNhalf, portionsFoodpackBACGNhalf);

            //Create Delivery PF
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryPFName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryPFName, customerPFName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio PF
                deliveryLoadingPage.AddService(servicioPFName);
                deliveryLoadingPage.AddQuantity("16");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackBACGN, false, null, null, portionsFoodpackBACGN);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackBACGNhalf, false, null, null, portionsFoodpackBACGNhalf);
                deliveryLoadingPage.CloseFoodPackagingModal();
                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryPFName);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(deliveryPFName), "La delivery {0} n'a pas été crée", deliveryPFName);

            //Create delivery round PF
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundPF);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRoundPF, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryPFName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundPF);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(deliveryRoundPF), "Le delivery round {0} n'a pas été créé.", deliveryRoundPF);

            //DISPATCH DELIVERY WITH SERVICIO PF = 16
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, servicioPFName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(servicioPFName)) > 0, "Le service {0} n'a pas de menu associé.", servicioPFName);
        }

        //Create Datas Delivery Condren (packaging by recipe variant, 2 recipes)
        [TestMethod]
        [Priority(14)]
        [Timeout(Timeout)]
        public void PR_SETUP_Create_Datas_MethodByRecipeVariant2Recipes()
        {
            // Prepare
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeType2 = TestContext.Properties["Production_RecipeType2"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();

            // Act
            //Create Customer Condren
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, condrenCustomerName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(condrenCustomerName, condrenCustomerCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, condrenCustomerName);
            }

            Assert.IsTrue(customersPage.GetFirstCustomerName().Equals(condrenCustomerName), "Le customer n'a pas été créé.");

            //Create SERVICIO Condren
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, condrenServicioName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(condrenServicioName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, condrenCustomerName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, condrenCustomerName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, condrenServicioName);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(condrenServicioName), "Le service " + condrenServicioName + " n'existe pas.");

            //Create Receta Condren
            //Receta Condren 1 recipe type = PLATO CALIENTE variante Adulto - almuerzo - poids de portion = 150g
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, condrenReceta1Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(condrenReceta1Name, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);
                recipeVariantPage.ClickOnFirstIngredient();
                recipeVariantPage.SetYield(yieldPF95);
                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);
                recipeVariantPage.ClickOnFirstIngredient();
                recipeVariantPage.SetYield(yieldPF95);
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, condrenReceta1Name);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(condrenReceta1Name), "La recette " + condrenReceta1Name + " n'existe pas.");

            //Receta Condren 2 recipe type = ENSALADA variante Adulto - almuerzo - poides de portion = 100g
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, condrenReceta2Name);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(condrenReceta2Name, recipeType2, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                // 1 variante pour ACE
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion100g);
                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion100g);
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, condrenReceta2Name);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(condrenReceta2Name), "La recette " + condrenReceta2Name + " n'existe pas.");

            //Create menu Condren RECETA CONDREN 1 methode %  coef 100% + RECETA CONDREN 2 methode std coef 2
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, condrenMenuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(condrenServicioName);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(condrenMenuName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(condrenReceta2Name);
                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeMethod("Std");
                menuDayViewPage.SetRecipeCoef("2");
                menuDayViewPage.AddRecipe(condrenReceta1Name);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsThisRecipeDisplayed(condrenReceta2Name))
                {
                    menuDayViewPage.AddRecipe(condrenReceta2Name);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeMethod("Std");
                    menuDayViewPage.SetRecipeCoef("2");
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(condrenReceta1Name))
                {
                    menuDayViewPage.AddRecipe(condrenReceta1Name);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsThisRecipeDisplayed(condrenReceta2Name))
                {
                    menuDayViewPage.AddRecipe(condrenReceta2Name);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeMethod("Std");
                    menuDayViewPage.SetRecipeCoef("2");
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(condrenReceta1Name))
                {
                    menuDayViewPage.AddRecipe(condrenReceta1Name);
                }

                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, condrenMenuName);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(condrenMenuName), "Le menu n'a pas été créé.");

            var parametersSetup = homePage.GoToParameters_SetupPage();
            parametersSetup.GoToTab_PAXPerRecipeVariant();
            parametersSetup.Filter(ParametersSetup.FilterType.SearchRecipe, condrenReceta1Name);
            parametersSetup.SelectAll();
            parametersSetup.ClickOnMultipleUpdate();
            parametersSetup.RemoveFoodPackPerRecipe(foodpackBACGN);
            parametersSetup.RemoveFoodPackPerRecipe(foodpackBACGNhalf);
            parametersSetup.AddFoodPackPerRecipe(foodpackBACGN, portionsFoodpackBACGN);
            parametersSetup.SelectAll();
            parametersSetup.ClickOnMultipleUpdate();
            parametersSetup.AddFoodPackPerRecipe(foodpackBACGNhalf, portionsFoodpackBACGNhalf);

            //Create Delivery CONDREN
            //BY RECIPE VARIANT, recipe types = PLATO CALIENTE, packaging name = BAC G / N & packaging name = BAC G / N 1 / 2
            //PAX PER PACKAGING 2 recipe type = ENSALADA  et BOCADILLO/ SANDWICH, packaging name = B1 - nb of portion = 1

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, condrenDeliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(condrenDeliveryName, condrenCustomerName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio CONDREN
                deliveryLoadingPage.AddService(condrenServicioName);
                deliveryLoadingPage.AddQuantity("15");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGN, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGNhalf, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodPaxPerPackaging, foodpackB1, false, null, null, portionsFoodpackB1, recipeType2);
                deliveryLoadingPage.CloseFoodPackagingModal();
                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, condrenDeliveryName);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(condrenDeliveryName), "La delivery n'a pas été crée");

            //Create delivery round CONDREN
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, condrenDeliveryRound);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(condrenDeliveryRound, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(condrenDeliveryName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, condrenDeliveryRound);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(condrenDeliveryRound), "Le delivery round n'a pas été créé.");

            //DISPATCH DELIVERY WITH SERVICIO CONDREN = 15
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, condrenServicioName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(condrenServicioName)) > 0, "Le service {0} n'a pas de menu associé.", condrenServicioName);
        }

        //Create Datas Delivery Barentin 
        [TestMethod]
        [Priority(15)]
        [Timeout(900000)]
        public void PR_SETUP_Create_Datas_MethodRecipeVariantAndGroupedByRecipe()
        {
            // Prepare            
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeVariantCollegioForACE = TestContext.Properties["Production_RecipeVariant2ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();
            string menuVariantCollegioForACE = TestContext.Properties["Production_MenuVariant2ForACE"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteACE);
            string siteIdACE = sitePage.CollectNewSiteID();

            // Act
            //Create Customer TLS
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, barentinCustomerName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(barentinCustomerName, barentinCustomerCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, barentinCustomerName);
            }

            //Create SERVICIO BARENTIN ADULTO
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, barentinServicioAdulto);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(barentinServicioAdulto, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, barentinCustomerName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, barentinCustomerName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, barentinServicioAdulto);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(barentinServicioAdulto), "Le service " + barentinServicioAdulto + " n'existe pas.");

            //Create SERVICIO BARENTIN COLLEGIO
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, barentinServicioCollegio);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(barentinServicioCollegio, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, barentinCustomerName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, barentinCustomerName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, barentinServicioCollegio);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(barentinServicioCollegio), "Le service " + barentinServicioCollegio + " n'existe pas.");

            //Create Receta Barentin 1
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinReceta1);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(barentinReceta1, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantCollegioForACE);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                recipeVariantPage2.AddIngredient(itemTest);
                recipeVariantPage2.SetTotalPortion(portion100g);
                recipesPage = recipeVariantPage2.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();

                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                if (!recipeVariantPage2.IsIngredientDisplayed())
                {
                    recipeVariantPage2.AddIngredient(itemTest);
                }
                recipeVariantPage2.SetTotalPortion(portion100g);

                recipesPage = recipeVariantPage2.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinReceta1);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(barentinReceta1), "La recette " + barentinReceta1 + " n'existe pas.");

            //Create Receta Barentin 2
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinRecetaNP2);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(barentinRecetaNP2, recipeTypeNP, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantCollegioForACE);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                recipeVariantPage2.AddIngredient(itemTest);
                recipeVariantPage2.SetTotalPortion(portion100g);
                recipesPage = recipeVariantPage2.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();

                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                if (!recipeVariantPage2.IsIngredientDisplayed())
                {
                    recipeVariantPage2.AddIngredient(itemTest);
                }
                recipeVariantPage2.SetTotalPortion(portion100g);

                recipesPage = recipeVariantPage2.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinRecetaNP2);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(barentinRecetaNP2), "La recette " + barentinReceta1 + " n'existe pas.");

            //Create Receta Barentin 3
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinRecetaBulk3);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(barentinRecetaBulk3, recipeTypeBulk, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantCollegioForACE);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemTest);
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                recipeVariantPage2.AddIngredient(itemTest);
                recipeVariantPage2.SetTotalPortion(portion100g);
                recipesPage = recipeVariantPage2.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();

                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemTest);
                }
                recipeVariantPage.SetTotalPortion(portion150g);

                var recipeVariantPage2 = recipeGeneralInfosPage.SelectVariant(recipeVariantCollegioForACE);
                if (!recipeVariantPage2.IsIngredientDisplayed())
                {
                    recipeVariantPage2.AddIngredient(itemTest);
                }
                recipeVariantPage2.SetTotalPortion(portion100g);

                recipesPage = recipeVariantPage2.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, barentinRecetaBulk3);
            Assert.IsTrue(recipesPage.GetFirstRecipeName().Contains(barentinRecetaBulk3), "La recette " + barentinRecetaBulk3 + " n'existe pas.");

            //Create Menu barentin adulto
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, barentinMenuAdulto);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(barentinServicioAdulto);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(barentinMenuAdulto, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(barentinReceta1);
                menuDayViewPage.AddRecipe(barentinRecetaNP2);
                menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinReceta1))
                {
                    menuDayViewPage.AddRecipe(barentinReceta1);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaNP2))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaNP2);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaBulk3))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinReceta1))
                {
                    menuDayViewPage.AddRecipe(barentinReceta1);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaNP2))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaNP2);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaBulk3))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, barentinMenuAdulto);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(barentinMenuAdulto), "Le menu n'a pas été créé.");

            //Create Menu barentin collegio
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, barentinMenuCollegio);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(barentinServicioCollegio);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(barentinMenuCollegio, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantCollegioForACE);
                menuDayViewPage.AddRecipe(barentinReceta1);
                menuDayViewPage.AddRecipe(barentinRecetaNP2);
                menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                //var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                //menuGeneralInformationPage.SetEndDate(DateUtils.Now.AddMonths(1));
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinReceta1))
                {
                    menuDayViewPage.AddRecipe(barentinReceta1);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaNP2))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaNP2);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaBulk3))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinReceta1))
                {
                    menuDayViewPage.AddRecipe(barentinReceta1);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaNP2))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaNP2);
                }
                if (!menuDayViewPage.IsThisRecipeDisplayed(barentinRecetaBulk3))
                {
                    menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                }
                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, barentinMenuCollegio);
            Assert.IsTrue(menusPage.GetFirstMenuName().Contains(barentinMenuCollegio), "Le menu n'a pas été créé.");

            var parametersSetup = homePage.GoToParameters_SetupPage();
            parametersSetup.GoToTab_PAXPerRecipeVariant();

            parametersSetup.AddPaxPerRecipe(barentinReceta1, menuVariantForACE, foodpackBACGN, portionsFoodpackBACGN, false);
            parametersSetup.AddPaxPerRecipe(barentinReceta1, menuVariantForACE, foodpackBACGNhalf, portionsFoodpackBACGNhalf, false);

            parametersSetup.AddPaxPerRecipe(barentinReceta1, menuVariantCollegioForACE, foodpackBACGN, "16", false);
            parametersSetup.AddPaxPerRecipe(barentinReceta1, menuVariantCollegioForACE, foodpackBACGNhalf, "8", false);

            parametersSetup.AddPaxPerRecipe(barentinRecetaBulk3, menuVariantForACE, foodpackEntier, "6", true);
            parametersSetup.AddPaxPerRecipe(barentinRecetaBulk3, menuVariantForACE, foodpackDemi, "3", true);

            parametersSetup.AddPaxPerRecipe(barentinRecetaBulk3, menuVariantCollegioForACE, foodpackEntier, "8", true);
            parametersSetup.AddPaxPerRecipe(barentinRecetaBulk3, menuVariantCollegioForACE, foodpackDemi, "4", true);

            parametersSetup.AddPaxPerRecipe(barentinRecetaNP2, menuVariantForACE, foodpackCRT, "50");
            parametersSetup.AddPaxPerRecipe(barentinRecetaNP2, menuVariantCollegioForACE, foodpackCRT, "50");

            //Settings/Set Up/ ajout PAX Per Recipe sur la recipe Bulk
            parametersSetup.GoToTab_PAXPerRecipe();
            parametersSetup.AddPaxPerRecipe(barentinRecetaNP2, null, foodpackCRT, "50");
            parametersSetup.AddPaxPerRecipe(barentinRecetaNP2, null, foodpackUnit, "1");
            parametersSetup.AddPaxPerRecipe(barentinRecetaNP2, null, foodpackSEAU, "10");

            //Create Delivery barentin
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, barentinDeliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(barentinDeliveryName, barentinCustomerName, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                // add servicio adulto
                deliveryLoadingPage.AddService(barentinServicioAdulto);
                deliveryLoadingPage.AddQuantity("10");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGN, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGNhalf, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackEntier, false, null, null, null, recipeTypeBulk);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackDemi, false, null, null, null, recipeTypeBulk);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackCRT, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackSEAU, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackUnit, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.CloseFoodPackagingModal();

                // add servicio collegio
                deliveryLoadingPage.AddService(barentinServicioCollegio);
                deliveryLoadingPage.AddQuantity("50");
                deliveryLoadingPage.AddPackaging();
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGN, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackBACGNhalf, false, null, null, null, recipeType1);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackEntier, false, null, null, null, recipeTypeBulk);
                deliveryLoadingPage.FillField_FoodPackaging(methodByRecipeVariant, foodpackDemi, false, null, null, null, recipeTypeBulk);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackCRT, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackSEAU, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.FillField_FoodPackaging(methodGroupedByRecipe, foodpackUnit, false, null, null, null, recipeTypeNP);
                deliveryLoadingPage.CloseFoodPackagingModal();

                deliveryLoadingPage.BackToList();
                deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, barentinDeliveryName);
            }

            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(barentinDeliveryName), "La delivery n'a pas été crée");

            //Create delivery round barentin
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, barentinDeliveryRound);

            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(barentinDeliveryRound, siteIdACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(barentinDeliveryName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, barentinDeliveryRound);
            Assert.IsTrue(deliveryRoundPage.GetFirstDeliveryRound().Contains(barentinDeliveryRound), "Le delivery round n'a pas été créé.");

            //Dispatch DELIVERY WITH SERVICIO Barentin adulto (10) et collegio (50)
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, "BARENTIN");
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(barentinServicioAdulto)) > 0, "Le service {0} n'a pas de menu associé.", barentinServicioAdulto);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(barentinServicioCollegio)) > 0, "Le service {0} n'a pas de menu associé.", barentinServicioCollegio);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Filter_ByDeliveryRoundName()
        {
            //Prepare
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, yesterday);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRound3Name);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var deliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenuNameAndQtyForOneDeliveryRound();
            Assert.IsTrue(deliveryRoundRecipeNameQties.DeliveryRoundName.Contains(deliveryRound3Name), "Le filtre SEARCH n'a pas été appliqué correctement");

            //Act
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, yesterday);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRound2Name);
            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            deliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenuNameAndQtyForOneDeliveryRound();

            Assert.IsTrue(deliveryRoundRecipeNameQties.DeliveryRoundName.Contains(deliveryRound2Name), "Le filtre SEARCH n'a pas été appliqué correctement");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Filter_ByDate()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string deliveryRoundNameToTest = deliveryRound2Name;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, yesterday);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundNameToTest);

            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var deliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenuNameAndQtyForOneDeliveryRound();

            // Go to Production - Dispatch
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryRoundRecipeNameQties.DeliveryRoundName);
            var quantityToProduceTab = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            Assert.IsFalse(quantityToProduceTab.CanUpdateQty(), "Cette date n'est pas valide puisque les quantités to produce du dispatch sont toujours modifiables.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Filter_BySite()
        {
            //Prepare
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();


            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, yesterday);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var allDeliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenusNamesAndQtyForAllDeliveryRound();

            // Go to Customer - Delivery round
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();

            foreach (var deliveryRound in allDeliveryRoundRecipeNameQties)
            {
                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRound.Value.DeliveryRoundName);
                bool verifySite = deliveryRoundPage.VerifySite(siteMAD);
                Assert.IsTrue(verifySite, MessageErreur.FILTRE_ERRONE, "Sites");
            }
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Filter_ByRecipeName()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, yesterday);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            //recipe1Name
            setupPage.Filter(SetupPage.FilterType.RecipeName, recipe2Name);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var deliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenuNameAndQtyForOneDeliveryRound();
            Assert.IsTrue(deliveryRoundRecipeNameQties.RecipeName.Contains(recipe2Name), MessageErreur.FILTRE_ERRONE, "Search by Recipe name");

            //recipe3Name
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, yesterday);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.RecipeName, recipe3Name);
            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            deliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenuNameAndQtyForOneDeliveryRound();
            Assert.IsTrue(deliveryRoundRecipeNameQties.RecipeName.Contains(recipe3Name), MessageErreur.FILTRE_ERRONE, "Search by Recipe name");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Filter_ByCustomerTypes()
        {
            //Prepare
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string customerType2 = TestContext.Properties["Production_CustomerType2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            //customer type Colectividades
            setupPage.Filter(SetupPage.FilterType.CustomerTypes, customerType1);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var allDeliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenusNamesAndQtyForAllDeliveryRound();

            // Go to Customer - Customer
            var customerPage = homePage.GoToCustomers_CustomerPage();

            foreach (var deliveryRound in allDeliveryRoundRecipeNameQties)
            {
                customerPage.ResetFilters();
                customerPage.Filter(CustomerPage.FilterType.Search, deliveryRound.Value.CustomerName);
                Assert.IsTrue(customerPage.VerifyTypeCustomer(customerType1), MessageErreur.FILTRE_ERRONE, "Customer Types");
            }

            //customer type Financiero
            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            setupPage = setupFilterAndFavoritesPage.Validate();

            //customer type Colectividades
            setupPage.Filter(SetupPage.FilterType.CustomerTypes, customerType2);
            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            allDeliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenusNamesAndQtyForAllDeliveryRound();

            // Go to Customer - Customer
            customerPage = homePage.GoToCustomers_CustomerPage();

            foreach (var deliveryRound in allDeliveryRoundRecipeNameQties)
            {
                customerPage.ResetFilters();
                customerPage.Filter(CustomerPage.FilterType.Search, deliveryRound.Value.CustomerName);
                Assert.IsTrue(customerPage.VerifyTypeCustomer(customerType2), MessageErreur.FILTRE_ERRONE, "Customer Types");
            }
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Filter_ByCustomers()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.Customers, customer3Name);

            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var allDeliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenusNamesAndQtyForAllDeliveryRound();

            foreach (var deliveryRound in allDeliveryRoundRecipeNameQties)
            {
                Assert.IsTrue(deliveryRound.Value.CustomerName.Equals(customer3Name), MessageErreur.FILTRE_ERRONE, "Customers");
            }
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Filter_ByDeliveries()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, "PR MAD");
            if (!setupPage.IsResultsDisplayed())
            {
                setupPage.Filter(SetupPage.FilterType.StartDate, yesterday);
            }
            setupPage.Filter(SetupPage.FilterType.Deliveries, delivery3Name);

            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var allDeliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenusNamesAndQtyForAllDeliveryRound();

            foreach (var deliveryRound in allDeliveryRoundRecipeNameQties)
            {
                Assert.IsTrue(deliveryRound.Value.DeliveryName.Equals(delivery3Name), MessageErreur.FILTRE_ERRONE, "Deliveries");
            }
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Filter_ByMealTypes()
        {
            //Prepare
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            string mealVariant1 = TestContext.Properties["Production_Meal1"].ToString();
            string mealVariant2 = TestContext.Properties["Production_Meal2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            //meal type 1
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, "PR MAD");
            if (!setupPage.IsResultsDisplayed())
            {
                setupPage.Filter(SetupPage.FilterType.StartDate, yesterday);
            }
            setupPage.Filter(SetupPage.FilterType.MealTypes, mealVariant1);

            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var allDeliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenusNamesAndQtyForAllDeliveryRound();

            foreach (var deliveryRound in allDeliveryRoundRecipeNameQties)
            {
                Assert.IsTrue(deliveryRound.Value.MenuVariant.Contains(mealVariant1), MessageErreur.FILTRE_ERRONE, "Meal Types");
            }

            //meal type 2
            setupPage.Filter(SetupPage.FilterType.MealTypes, mealVariant2);

            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            allDeliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenusNamesAndQtyForAllDeliveryRound();

            foreach (var deliveryRound in allDeliveryRoundRecipeNameQties)
            {
                Assert.IsTrue(deliveryRound.Value.MenuVariant.Contains(mealVariant2), MessageErreur.FILTRE_ERRONE, "Meal Types");
            }
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Filter_ByWorkshops()
        {
            //Prepare
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            string workshop1 = TestContext.Properties["Production_Workshop1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            //meal type 1
            setupPage.Filter(SetupPage.FilterType.Sites, siteMAD);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, "PR MAD");
            if (!setupPage.IsResultsDisplayed())
            {
                setupPage.Filter(SetupPage.FilterType.StartDate, yesterday);
            }
            setupPage.Filter(SetupPage.FilterType.Workshops, workshop1);

            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var allDeliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenusNamesAndQtyForAllDeliveryRound();

            foreach (var deliveryRound in allDeliveryRoundRecipeNameQties)
            {
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, deliveryRound.Value.RecipeName);
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                recipeGeneralInfosPage.ClickOnEditInformation();

                Assert.IsTrue(recipeGeneralInfosPage.GetWorkshop().Equals(workshop1), MessageErreur.FILTRE_ERRONE, "Meal Types");
            }
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Unfold_All_And_Fold_All()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            // Arrange
            HomePage HomePage= LogInAsAdmin();
           
            //Act
            var setupFilterAndFavoritesPage = HomePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();

            //Unfold
            bool unfold = setupDeliveryRoundTabPage.FoldUnfoldAll();
            Assert.IsTrue(unfold, "Les éléments du tableau ne sont pas dépliés");
            //Fold
            Assert.IsTrue(!unfold, "Les éléments du tableau ne sont pas repliés");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Link_Recipe()
        {
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteMAD);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRound3Name);

            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var deliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenuNameAndQtyForOneDeliveryRound();

            // Check if total is the quantity to produce by the weight per pack
            Assert.AreEqual((deliveryRoundRecipeNameQties.QtyToProduce * deliveryRoundRecipeNameQties.WeightPerPack), deliveryRoundRecipeNameQties.Total, "Le total ne correspond pas à la 'quantity to produce' et au 'weight per pack'0 .");

            var recipeVariantPage = setupDeliveryRoundTabPage.EditRecipe();

            string recipeName = recipeVariantPage.GetRecipeName();

            Assert.IsTrue(recipeName.Contains(deliveryRoundRecipeNameQties.RecipeName.ToUpper()), "La page d'édition de la recette donnée n'est pas visible.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Print_DeliveryRound()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            SetupPage setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            PrintReportPage printReportPage = setupPage.Print(SetupPage.PrintType.DeliveryRound, newVersionPrint);
            bool isReportGenerated = printReportPage.IsReportGenerated();
            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF Print SetUp n'a pas pu être généré par l'application.");
            //Pour delivery perpaxgrouped DR GR TLS, c'est voulu de laisser séparer dans le print et de ne pas rassembler les PAX
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Print_DeliveryRoundByRecipes()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();

            setupPage.Filter(SetupPage.FilterType.Sites, siteMAD);

            var printReportPage = setupPage.Print(SetupPage.PrintType.DeliveryRoundByRecipes, newVersionPrint);

            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF Print SetUp n'a pas pu être généré par l'application.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Print_DeliveryNoteByRecipes()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            var printReportPage = setupPage.Print(SetupPage.PrintType.DeliveryNoteByRecipes, newVersionPrint);

            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF Print SetUp n'a pas pu être généré par l'application.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Print_DeliveryNoteByServices()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();

            setupPage.Filter(SetupPage.FilterType.Sites, siteMAD);

            var printReportPage = setupPage.Print(SetupPage.PrintType.DeliveryNoteByServices, newVersionPrint);

            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF Print SetUp n'a pas pu être généré par l'application.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Print_DeliveryNoteValorized()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            var printReportPage = setupPage.Print(SetupPage.PrintType.DeliveryNoteValorized, newVersionPrint);

            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF Print SetUp n'a pas pu être généré par l'application.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Print_FoodPackReport()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.ClearDownloads();
            var printReportPage = setupPage.Print(SetupPage.PrintType.FoodPackReport, newVersionPrint);
            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF Print SetUp n'a pas pu être généré par l'application.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Print_FoodPackGroupByDelivery()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.ClearDownloads();
            var printReportPage = setupPage.Print(SetupPage.PrintType.FoodPackGroupByDelivery, newVersionPrint);
            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF Print SetUp n'a pas pu être généré par l'application.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_ExportCSVFile()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRound2Name);

            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var deliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenuNameAndQtyForOneDeliveryRound();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            var siteName = setupPage.ReadCSVFile(filePath, "Site name");
            Assert.IsTrue(siteACE.Equals(siteName), "Sur le fichier csv, le site ne correspond pas.");

            var recipeName = setupPage.ReadCSVFile(filePath, "Recipe name");
            var isRecipeName = deliveryRoundRecipeNameQties.RecipeName.Equals(recipeName);
            Assert.IsTrue(isRecipeName, "Sur le fichier csv, le nom du menu ne correspond pas.");

            var guestType = setupPage.ReadCSVFile(filePath, "Guest type");
            var isGuestType = deliveryRoundRecipeNameQties.MenuVariant.Contains(guestType);
            Assert.IsTrue(isGuestType, "Sur le fichier csv, le nom du guest type ne correspond pas.");

            var mealType = setupPage.ReadCSVFile(filePath, "Meal type");
            var isMealType = deliveryRoundRecipeNameQties.MenuVariant.Contains(mealType);
            Assert.IsTrue(isMealType, "Sur le fichier csv, le nom du meal type ne correspond pas.");

            var customerName = setupPage.ReadCSVFile(filePath, "Customer name");
            var isCustomerName = deliveryRoundRecipeNameQties.CustomerName.Contains(customerName);
            Assert.IsTrue(isCustomerName, "Sur le fichier csv, le nom du Customer name ne correspond pas.");

            var deliveryName = setupPage.ReadCSVFile(filePath, "Delivery name");
            var isDeliveryName = deliveryRoundRecipeNameQties.DeliveryName.Contains(deliveryName);
            Assert.IsTrue(isDeliveryName, "Sur le fichier csv, le nom du Delivery name ne correspond pas.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_GoToSettings()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            SetupPage setupPage = setupFilterAndFavoritesPage.Validate();
            ParametersSetup parametersSetupPage = setupPage.GoToParameters_SetupSettingsFromSetupPage();
            bool isSelected = parametersSetupPage.CheckSetupTabSelected();
            Assert.IsTrue(isSelected, "Le lien Settings ne renvoit pas sur la pas Parameters Setup");
        }


        //TESTS v2

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckPAX_MethodIndividualMultiportion()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string ESINDMULServiceQuantity1 = "1";
            string ESINDMULServiceQuantity2 = "10";
            string ESINDServiceQuantity = "5";
            string ESMULServiceQuantity = "10";
            var date = DateUtils.Now.AddYears(-1);
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //DELIVERY WITH SERVICIO ES IND = 5, SERVICIO ES MUL = 10 & SERVICIO INDMUL = 1
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, ESDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDServicioName, ESINDServiceQuantity);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESMULServicioName, ESMULServiceQuantity);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDMULServicioName, ESINDMULServiceQuantity1);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isDispatchValidated = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isDispatchValidated, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESMULServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDMULServicioName);

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundES);
            setupPage.Filter(SetupPage.FilterType.StartDate, date);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service ES IND MUL SERVICIO
            var recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(deliveryRoundES, ESINDMULServicioName)[0];
            Assert.IsTrue(recipeNameDisplayed.Contains(recetaESName), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", ESINDMULServicioName, deliveryRoundES, recipeNameDisplayed, recetaESName);
            var PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(deliveryRoundES, ESINDMULServicioName);
            Assert.IsTrue(PAXDisplayed[0].Contains(ESINDMULServiceQuantity1), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <PAX : {3}>", ESINDMULServicioName, deliveryRoundES, PAXDisplayed[0], ESINDMULServiceQuantity1);
            //Service ES IND SERVICIO
            recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(deliveryRoundES, ESINDServicioName)[0];
            Assert.IsTrue(recipeNameDisplayed.Contains(recetaESName), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", ESINDServicioName, deliveryRoundES, recipeNameDisplayed, recetaESName);
            PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(deliveryRoundES, ESINDServicioName);
            Assert.IsTrue(PAXDisplayed[0].Contains(ESINDServiceQuantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <PAX : {3}>", ESINDServicioName, deliveryRoundES, PAXDisplayed[0], ESINDServiceQuantity);
            //Service ES MUL SERVICIO
            recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(deliveryRoundES, ESMULServicioName)[0];
            Assert.IsTrue(recipeNameDisplayed.Contains(recetaESName), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", ESMULServicioName, deliveryRoundES, recipeNameDisplayed, recetaESName);
            PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(deliveryRoundES, ESMULServicioName);
            Assert.IsTrue(PAXDisplayed[0].Contains(ESMULServiceQuantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <PAX : {3}>", ESMULServicioName, deliveryRoundES, PAXDisplayed[0], ESMULServiceQuantity);

            //DELIVERY WITH SERVICIO ES IND = 5, SERVICIO ES MUL = 10 & SERVICIO INDMUL = 10
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, ESDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDServicioName, ESINDServiceQuantity);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESMULServicioName, ESMULServiceQuantity);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDMULServicioName, ESINDMULServiceQuantity2);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isDispatchValidatedByColorDay = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isDispatchValidatedByColorDay, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESMULServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDMULServicioName);

            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundES);
            setupPage.Filter(SetupPage.FilterType.StartDate, date);
            setupDeliveryRoundTabPage.WaitPageLoading();
            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service ES IND MUL SERVICIO
            recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(deliveryRoundES, ESINDMULServicioName)[0];
            Assert.IsTrue(recipeNameDisplayed.Contains(recetaESName), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", ESINDMULServicioName, deliveryRoundES, recipeNameDisplayed, recetaESName);
            PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(deliveryRoundES, ESINDMULServicioName);
            Assert.IsTrue(PAXDisplayed[0].Contains(ESINDMULServiceQuantity2), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <PAX : {3}>", ESINDMULServicioName, deliveryRoundES, PAXDisplayed[0], ESINDMULServiceQuantity2);
            //Service ES IND SERVICIO
            recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(deliveryRoundES, ESINDServicioName)[0];
            Assert.IsTrue(recipeNameDisplayed.Contains(recetaESName), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", ESINDServicioName, deliveryRoundES, recipeNameDisplayed, recetaESName);
            PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(deliveryRoundES, ESINDServicioName);
            Assert.IsTrue(PAXDisplayed[0].Contains(ESINDServiceQuantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <PAX : {3}>", ESINDServicioName, deliveryRoundES, PAXDisplayed[0], ESINDServiceQuantity);
            //Service ES MUL SERVICIO
            recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(deliveryRoundES, ESMULServicioName)[0];
            Assert.IsTrue(recipeNameDisplayed.Contains(recetaESName), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", ESMULServicioName, deliveryRoundES, recipeNameDisplayed, recetaESName);
            PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(deliveryRoundES, ESMULServicioName);
            Assert.IsTrue(PAXDisplayed[0].Contains(ESMULServiceQuantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <PAX : {3}>", ESMULServicioName, deliveryRoundES, PAXDisplayed[0], ESMULServiceQuantity);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckPAX_MethodPaxPerPackaging()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string TLSServiceQuantity1 = "10";
            string TLSServiceQuantity2 = "12";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //DELIVERY WITH SERVICIO TLS = 10
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, TLSServicioName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, TLSServiceQuantity1);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, TLSDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service TLS SERVICIO
            var recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(TLSDeliveryRound, TLSServicioName);
            Assert.IsTrue(recipeNameDisplayed.Contains(TLSRecetaName), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", TLSServicioName, TLSDeliveryRound, recipeNameDisplayed, TLSRecetaName);
            var PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(TLSDeliveryRound, TLSServicioName);
            Assert.IsTrue(PAXDisplayed[0].Contains(TLSServiceQuantity1), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <PAX : {3}>", TLSServicioName, TLSDeliveryRound, PAXDisplayed[0], TLSServiceQuantity1);

            //DELIVERY WITH SERVICIO TLS = 12
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, TLSServicioName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, TLSServiceQuantity2);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, TLSDeliveryRound);
            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();
            //Service TLS SERVICIO
            recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(TLSDeliveryRound, TLSServicioName);
            Assert.IsTrue(recipeNameDisplayed.Contains(TLSRecetaName), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", TLSServicioName, TLSDeliveryRound, recipeNameDisplayed, TLSRecetaName);
            PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(TLSDeliveryRound, TLSServicioName);
            Assert.IsTrue(PAXDisplayed[0].Contains(TLSServiceQuantity2), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <PAX : {3}>", TLSServicioName, TLSDeliveryRound, PAXDisplayed[0], TLSServiceQuantity2);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckPAX_MethodPaxPerPackagingGrouped()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string TLSGRPAXTotalQuantity = "15";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundTLSGR);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            var recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(deliveryRoundTLSGR, "grouped by recipe");
            Assert.IsTrue(recipeNameDisplayed.Contains(recetaTLSGRName), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", servicioTLSGR1Name, deliveryRoundTLSGR, recipeNameDisplayed, recetaTLSGRName);
            var PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(deliveryRoundTLSGR, "grouped by recipe");
            Assert.IsTrue(PAXDisplayed[0].Contains(TLSGRPAXTotalQuantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <PAX : {3}>", "grouped by recipe", deliveryRoundTLSGR, PAXDisplayed[0], TLSGRPAXTotalQuantity);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckPAX_MethodByRecipeVariant()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string PFService2Quantity = "16";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundPF);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service PF SERVICIO
            var recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(deliveryRoundPF, servicioPFName);
            Assert.IsTrue(recipeNameDisplayed.Contains(recetaPFName), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", servicioPFName, deliveryRoundPF, recipeNameDisplayed, recetaPFName);
            var PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(deliveryRoundPF, servicioPFName);
            Assert.IsTrue(PAXDisplayed[0].Contains(PFService2Quantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <PAX : {3}>", servicioPFName, deliveryRoundPF, PAXDisplayed, PFService2Quantity);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckPAX_MethodByRecipeVariant2Recipes()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string CondrenServiceQuantity = "15";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, condrenDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service PF SERVICIO
            var recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(condrenDeliveryRound, condrenServicioName);
            //condren receta 1
            Assert.IsTrue(recipeNameDisplayed.Contains(condrenReceta1Name), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", condrenServicioName, condrenDeliveryRound, recipeNameDisplayed, condrenReceta1Name);
            Assert.IsTrue(recipeNameDisplayed.Contains(condrenReceta2Name), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", condrenServicioName, condrenDeliveryRound, recipeNameDisplayed, condrenReceta2Name);
            var PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(condrenDeliveryRound, condrenServicioName);
            Assert.IsTrue(PAXDisplayed[0].Contains(CondrenServiceQuantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0} et de la recette {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <PAX : {4}>", condrenServicioName, condrenReceta1Name, condrenDeliveryRound, PAXDisplayed[0], CondrenServiceQuantity);
            Assert.IsTrue(PAXDisplayed[1].Contains(CondrenServiceQuantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0} et de la recette {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <PAX : {4}>", condrenServicioName, condrenReceta2Name, condrenDeliveryRound, PAXDisplayed[1], CondrenServiceQuantity);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckPAX_MethodRecipeVariantAndGroupedByRecipe()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string barentinAdultoQuantity = "10";
            string barentinCollegioQuantity = "50";
            string barentinTotalQuantity = "60";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, barentinDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //bulk packaging
            var recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(barentinDeliveryRound, "bulk packaging");
            Assert.IsTrue(recipeNameDisplayed.Contains(barentinRecetaBulk3), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", "bulk packaging", barentinDeliveryRound, recipeNameDisplayed, barentinRecetaBulk3);
            var PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(barentinDeliveryRound, "bulk packaging");
            Assert.IsTrue(PAXDisplayed[0].Contains(barentinTotalQuantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0} et de la recette {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <PAX : {4}>", "bulk packaging", barentinRecetaBulk3, barentinDeliveryRound, PAXDisplayed[0], barentinTotalQuantity);

            //Grouped by
            recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(barentinDeliveryRound, "grouped by recipe");
            Assert.IsTrue(recipeNameDisplayed.Contains(barentinRecetaNP2), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", "grouped by recipe", barentinDeliveryRound, recipeNameDisplayed, barentinRecetaNP2);
            PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(barentinDeliveryRound, "grouped by recipe");
            Assert.IsTrue(PAXDisplayed[0].Contains(barentinTotalQuantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0} et de la recette {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <PAX : {4}>", "grouped by recipe", barentinRecetaNP2, barentinDeliveryRound, PAXDisplayed[0], barentinTotalQuantity);

            //SERVICIO BARENTIN ADULTO 
            recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(barentinDeliveryRound, barentinServicioAdulto);
            Assert.IsTrue(recipeNameDisplayed.Contains(barentinReceta1), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", barentinServicioAdulto, barentinDeliveryRound, recipeNameDisplayed, barentinReceta1);
            PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(barentinDeliveryRound, barentinServicioAdulto);
            Assert.IsTrue(PAXDisplayed[0].Contains(barentinAdultoQuantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0} et de la recette {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <PAX : {4}>", barentinServicioAdulto, barentinReceta1, barentinDeliveryRound, PAXDisplayed[0], barentinAdultoQuantity);

            //SERVICIO BARENTIN COLLEGIO
            recipeNameDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceRecipeName(barentinDeliveryRound, barentinServicioCollegio);
            Assert.IsTrue(recipeNameDisplayed.Contains(barentinReceta1), "Dans Setup, onglet Delivery Round, le nom de la recette du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <{3}>", barentinServicioCollegio, barentinDeliveryRound, recipeNameDisplayed, barentinReceta1);
            PAXDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServicePAX(barentinDeliveryRound, barentinServicioCollegio);
            Assert.IsTrue(PAXDisplayed[0].Contains(barentinCollegioQuantity), "Dans Setup, onglet Delivery Round, la quantité de PAX du service {0} et de la recette {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <PAX : {4}>", barentinServicioCollegio, barentinReceta1, barentinDeliveryRound, PAXDisplayed[0], barentinCollegioQuantity);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckQuantityToProduce_MethodIndividualMultiportion()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity5 = "5";
            string quantity10 = "10";
            string quantity1 = "1";

            string ESINDMULServiceQuantityToProduce1 = "2 x Individual";
            string ESINDMULServiceQuantityToProduce2 = "1 x Multiporcion";
            string ESINDServiceQuantityToProduce = "3 x Individual";
            string ESMULServiceQuantityToProduce = "1 x Multiporcion";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //DELIVERY WITH SERVICIO ES IND = 5, SERVICIO ES MUL = 10 & SERVICIO INDMUL = 1
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, ESDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDServicioName, quantity5);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESMULServicioName, quantity10);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDMULServicioName, quantity1);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isDispatchValidated = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isDispatchValidated, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESMULServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDMULServicioName);

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundES);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();

            //Service ES IND MUL SERVICIO
            var quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(deliveryRoundES, ESINDMULServicioName);
            Assert.IsTrue(quantityToProduce[0].Contains(ESINDMULServiceQuantityToProduce1), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", ESINDMULServicioName, deliveryRoundES, quantityToProduce[0], ESINDMULServiceQuantityToProduce1);
            //Service ES IND SERVICIO
            quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(deliveryRoundES, ESINDServicioName);
            Assert.IsTrue(quantityToProduce[0].Contains(ESINDServiceQuantityToProduce), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", ESINDServicioName, deliveryRoundES, quantityToProduce[0], ESINDServiceQuantityToProduce);
            //Service ES MUL SERVICIO
            quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(deliveryRoundES, ESMULServicioName);
            Assert.IsTrue(quantityToProduce[0].Contains(ESMULServiceQuantityToProduce), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", ESMULServicioName, deliveryRoundES, quantityToProduce[0], ESMULServiceQuantityToProduce);

            //DELIVERY WITH SERVICIO ES IND = 5, SERVICIO ES MUL = 10 & SERVICIO INDMUL = 10
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, ESDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDServicioName, quantity5);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESMULServicioName, quantity10);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDMULServicioName, quantity10);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isDispatchValidatedByColorDay = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isDispatchValidatedByColorDay, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESMULServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDMULServicioName);

            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundES);
            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();

            //Service ES IND MUL SERVICIO
            quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(deliveryRoundES, ESINDMULServicioName);
            Assert.IsTrue(quantityToProduce[0].Contains(ESINDMULServiceQuantityToProduce2), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", ESINDMULServicioName, deliveryRoundES, quantityToProduce[0], ESINDMULServiceQuantityToProduce2);

            //Service ES IND SERVICIO
            quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(deliveryRoundES, ESINDServicioName);
            Assert.IsTrue(quantityToProduce[0].Contains(ESINDServiceQuantityToProduce), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", ESINDServicioName, deliveryRoundES, quantityToProduce[0], ESINDServiceQuantityToProduce);

            //Service ES MUL SERVICIO
            quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(deliveryRoundES, ESMULServicioName);
            Assert.IsTrue(quantityToProduce[0].Contains(ESMULServiceQuantityToProduce), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", ESMULServicioName, deliveryRoundES, quantityToProduce[0], ESMULServiceQuantityToProduce);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckQuantityToProduce_MethodPaxPerPackaging()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string TLSServiceQuantity1 = "10";
            string TLSServiceQuantity2 = "12";
            string TLSServiceQuantityToProduce1 = "3 x B6 2 x B1";
            string TLSServiceQuantityToProduce2 = "4 x B6";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //DELIVERY WITH SERVICIO TLS = 10
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, TLSServicioName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, TLSServiceQuantity1);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isDispatchValidated = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isDispatchValidated, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, TLSDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service TLS SERVICIO
            var quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(TLSDeliveryRound, TLSServicioName);
            Assert.IsTrue(quantityToProduce[0].Contains(TLSServiceQuantityToProduce1), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", TLSServicioName, TLSDeliveryRound, quantityToProduce[0], TLSServiceQuantityToProduce1);

            //DELIVERY WITH SERVICIO TLS = 12
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, TLSServicioName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, TLSServiceQuantity2);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isDispatchValidatedByColorDay = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isDispatchValidatedByColorDay, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, TLSDeliveryRound);
            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service TLS SERVICIO
            quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(TLSDeliveryRound, TLSServicioName);
            Assert.IsTrue(quantityToProduce[0].Contains(TLSServiceQuantityToProduce2), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", TLSServicioName, TLSDeliveryRound, quantityToProduce[0], TLSServiceQuantityToProduce2);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckQuantityToProduce_MethodPaxPerPackagingGrouped()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string TLSGRTotalQuantityToProduce = "2 x B6 3 x B1";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundTLSGR);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service TLS GR 1 SERVICIO
            var quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(deliveryRoundTLSGR, "grouped by recipe");
            Assert.IsTrue(quantityToProduce[0].Contains(TLSGRTotalQuantityToProduce), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", "grouped by recipe", deliveryRoundTLSGR, quantityToProduce[0], TLSGRTotalQuantityToProduce);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckQuantityToProduce_MethodByRecipeVariant()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string PFServiceQuantityToProduce = "1 x BAC G/N 1 x BAC G/N 1/2 (+1) x";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundPF);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service PF SERVICIO
            var quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(deliveryRoundPF, servicioPFName);
            Assert.IsTrue(quantityToProduce[0].Contains(PFServiceQuantityToProduce), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", servicioPFName, deliveryRoundPF, quantityToProduce[0], PFServiceQuantityToProduce);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckQuantityToProduce_MethodByRecipeVariant2Recipes()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string CondrenServiceQuantityToProduce1 = "1 x BAC G/N 1 x BAC G/N 1/2";
            string CondrenServiceQuantityToProduce2 = "2 x B1";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, condrenDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service Condren SERVICIO - RECETA CONDREN 1
            var quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(condrenDeliveryRound, condrenServicioName);
            Assert.IsTrue(quantityToProduce[1].Contains(CondrenServiceQuantityToProduce1), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, associé à la rectte {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <Qty to produce: {4}>", condrenServicioName, condrenReceta1Name, condrenDeliveryRound, quantityToProduce[1], CondrenServiceQuantityToProduce1);

            //Service Condren SERVICIO - RECETA CONDREN 2
            quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(condrenDeliveryRound, condrenServicioName);
            Assert.IsTrue(quantityToProduce[0].Contains(CondrenServiceQuantityToProduce2), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, associé à la rectte {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <Qty to produce: {4}>", condrenServicioName, condrenReceta2Name, condrenDeliveryRound, quantityToProduce[0], CondrenServiceQuantityToProduce2);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckQuantityToProduce_MethodRecipeVariantAndGroupedByRecipe()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string barentinBulkPackagingQuantityToProduce = "7 x ENTIER bulk remainder";
            string barentinGroupedByQuantityToProduce = "1 x CRT 1 x SEAU";
            string barentinAdultoServicioQuantityToProduce = "1 x BAC G/N";
            string barentinCollegioServicioQuantityToProduce = "3 x BAC G/N (+2) x";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, barentinDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //bulk packaging
            var quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(barentinDeliveryRound, "bulk packaging");
            Assert.IsTrue(quantityToProduce[0].Contains(barentinBulkPackagingQuantityToProduce), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", "bulk packaging", barentinDeliveryRound, quantityToProduce[0], barentinBulkPackagingQuantityToProduce);

            //Grouped by
            quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(barentinDeliveryRound, "grouped by recipe");
            Assert.IsTrue(quantityToProduce[0].Contains(barentinGroupedByQuantityToProduce), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", "grouped by recipe", barentinDeliveryRound, quantityToProduce[0], barentinGroupedByQuantityToProduce);

            //SERVICIO BARENTIN ADULTO 
            quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(barentinDeliveryRound, barentinServicioAdulto);
            Assert.IsTrue(quantityToProduce[0].Contains(barentinAdultoServicioQuantityToProduce), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", barentinServicioAdulto, barentinDeliveryRound, quantityToProduce[0], barentinAdultoServicioQuantityToProduce);

            //SERVICIO BARENTIN COLLEGIO
            quantityToProduce = setupDeliveryRoundTabPage.GetDeliveryRoundServiceQuantityToProduce(barentinDeliveryRound, barentinServicioCollegio);
            Assert.IsTrue(quantityToProduce[0].Contains(barentinCollegioServicioQuantityToProduce), "Dans Setup, onglet Delivery Round, la quantité to produce du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Qty to produce: {3}>", barentinServicioCollegio, barentinDeliveryRound, quantityToProduce[0], barentinCollegioServicioQuantityToProduce);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckWeightPerPack_MethodIndividualMultiportion()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity5 = "5";
            string quantity10 = "10";
            string quantity1 = "1";

            string ESINDMULServiceWeight1 = "150 g";
            string ESINDMULServiceWeight2 = "2250 g";
            string ESINDServiceWeight = "150 g";
            string ESMULServiceWeight = "1500 g";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //DELIVERY WITH SERVICIO ES IND = 5, SERVICIO ES MUL = 10 & SERVICIO INDMUL = 1
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, ESDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDServicioName, quantity5);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESMULServicioName, quantity10);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDMULServicioName, quantity1);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESMULServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDMULServicioName);

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundES);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();

            //Service ES IND MUL SERVICIO
            var weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(deliveryRoundES, ESINDMULServicioName);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(ESINDMULServiceWeight1), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", ESINDMULServicioName, deliveryRoundES, weightPerPackDisplayed[0], ESINDMULServiceWeight1);
            //Service ES IND SERVICIO
            weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(deliveryRoundES, ESINDServicioName);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(ESINDServiceWeight), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", ESINDServicioName, deliveryRoundES, weightPerPackDisplayed[0], ESINDServiceWeight);
            //Service ES MUL SERVICIO
            weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(deliveryRoundES, ESMULServicioName);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(ESMULServiceWeight), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", ESMULServicioName, deliveryRoundES, weightPerPackDisplayed[0], ESMULServiceWeight);

            //DELIVERY WITH SERVICIO ES IND = 5, SERVICIO ES MUL = 10 & SERVICIO INDMUL = 10
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, ESDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDServicioName, quantity5);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESMULServicioName, quantity10);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDMULServicioName, quantity10);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESMULServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDMULServicioName);

            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundES);
            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service ES IND MUL SERVICIO
            weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(deliveryRoundES, ESINDMULServicioName);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(ESINDMULServiceWeight2), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", ESINDMULServicioName, deliveryRoundES, weightPerPackDisplayed[0], ESINDMULServiceWeight2);
            //Service ES IND SERVICIO
            weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(deliveryRoundES, ESINDServicioName);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(ESINDServiceWeight), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack :: {3}>", ESINDServicioName, deliveryRoundES, weightPerPackDisplayed[0], ESINDServiceWeight);
            //Service ES MUL SERVICIO
            weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(deliveryRoundES, ESMULServicioName);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(ESMULServiceWeight), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", ESMULServicioName, deliveryRoundES, weightPerPackDisplayed[0], ESMULServiceWeight);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckWeightPerPack_MethodPaxPerPackaging()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string TLSServiceWeight1 = "900 g 150 g";
            string TLSServiceWeight2 = "900 g";

            string TLSServiceQuantity1 = "10";
            string TLSServiceQuantity2 = "12";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //DELIVERY WITH SERVICIO TLS = 10
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, TLSServicioName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, TLSServiceQuantity1);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, TLSDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();

            //Service TLS SERVICIO
            var weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(TLSDeliveryRound, TLSServicioName);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(TLSServiceWeight1), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", TLSServicioName, TLSDeliveryRound, weightPerPackDisplayed[0], TLSServiceWeight1);

            //DELIVERY WITH SERVICIO TLS = 12
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, TLSServicioName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, TLSServiceQuantity2);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, TLSDeliveryRound);
            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();

            //Service TLS SERVICIO
            weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(TLSDeliveryRound, TLSServicioName);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(TLSServiceWeight2), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", TLSServicioName, TLSDeliveryRound, weightPerPackDisplayed[0], TLSServiceWeight2);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckWeightPerPack_MethodPaxPerPackagingGrouped()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string TLSGRService1Weight = "600 g 100 g";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundTLSGR);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service TLS GR 1 SERVICIO
            var weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(deliveryRoundTLSGR, "grouped by recipe");
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(TLSGRService1Weight), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", servicioTLSGR1Name, deliveryRoundTLSGR, weightPerPackDisplayed[0], "grouped by recipe");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckWeightPerPack_MethodByRecipeVariant()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string pfServiceWeight = "1500 g 750 g 150 g";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundPF);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service PF SERVICIO
            var weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(deliveryRoundPF, servicioPFName);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(pfServiceWeight), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", servicioPFName, deliveryRoundPF, weightPerPackDisplayed[0], pfServiceWeight);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckWeightPerPack_MethodByRecipeVariant2Recipes()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string CondrenServiceQuantityToProduce1 = "1500 g 750 g";
            string CondrenServiceQuantityToProduce2 = "100 g";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, condrenDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service Condren SERVICIO - RECETA CONDREN 1
            var weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(condrenDeliveryRound, condrenServicioName);
            Assert.IsTrue(weightPerPackDisplayed[1].Contains(CondrenServiceQuantityToProduce1), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, associé à la rectte {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <Weight per pack : {4}>", condrenServicioName, condrenReceta1Name, condrenDeliveryRound, weightPerPackDisplayed[1], CondrenServiceQuantityToProduce1);

            //Service Condren SERVICIO - RECETA CONDREN 2
            weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(condrenDeliveryRound, condrenServicioName);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(CondrenServiceQuantityToProduce2), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, associé à la rectte {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <Weight per pack : {4}>", condrenServicioName, condrenReceta2Name, condrenDeliveryRound, weightPerPackDisplayed[0], CondrenServiceQuantityToProduce2);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckWeightPerPack_MethodRecipeVariantAndGroupedByRecipe()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string barentinBulkPackagingQuantityToProduce = "900 g 200 g";
            string barentinGroupedByQuantityToProduce = "5000 g 1000 g";
            string barentinAdultoServicioQuantityToProduce = "1500 g";
            string barentinCollegioServicioQuantityToProduce = "1600 g 100 g";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, barentinDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //bulk packaging
            var weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(barentinDeliveryRound, "bulk packaging");
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(barentinBulkPackagingQuantityToProduce), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", "bulk packaging", barentinDeliveryRound, weightPerPackDisplayed[0], barentinBulkPackagingQuantityToProduce);

            //Grouped by
            weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(barentinDeliveryRound, "grouped by recipe");
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(barentinGroupedByQuantityToProduce), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", "grouped by recipe", barentinDeliveryRound, weightPerPackDisplayed[0], barentinGroupedByQuantityToProduce);

            //SERVICIO BARENTIN ADULTO 
            weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(barentinDeliveryRound, barentinServicioAdulto);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(barentinAdultoServicioQuantityToProduce), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", barentinServicioAdulto, barentinDeliveryRound, weightPerPackDisplayed[0], barentinAdultoServicioQuantityToProduce);

            //SERVICIO BARENTIN COLLEGIO
            weightPerPackDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceWeights(barentinDeliveryRound, barentinServicioCollegio);
            Assert.IsTrue(weightPerPackDisplayed[0].Contains(barentinCollegioServicioQuantityToProduce), "Dans Setup, onglet Delivery Round, le weight per pack du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Weight per pack : {3}>", barentinServicioCollegio, barentinDeliveryRound, weightPerPackDisplayed[0], barentinCollegioServicioQuantityToProduce);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckTotal_MethodIndividualMultiportion()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity5 = "5";
            string quantity10 = "10";
            string quantity1 = "1";

            string ESINDMULServiceTotal1 = "300 g";
            string ESINDMULServiceTotal2 = "2250 g";
            string ESINDServiceTotal = "450 g";
            string ESMULServiceTotal = "1500 g";


            // Arrange
            HomePage homePage = LogInAsAdmin();

            //DELIVERY WITH SERVICIO ES IND = 5, SERVICIO ES MUL = 10 & SERVICIO INDMUL = 1
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, ESDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDServicioName, quantity5);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESMULServicioName, quantity10);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDMULServicioName, quantity1);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isDipatchValidatedByColorDay = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isDipatchValidatedByColorDay, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESMULServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDMULServicioName);

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundES);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service ES IND MUL SERVICIO
            var total = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(deliveryRoundES, ESINDMULServicioName);
            var actualTotal = total[0].Replace("Total : ", "").Trim();
            Assert.AreEqual(ESINDMULServiceTotal1, actualTotal, "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", ESINDMULServicioName, deliveryRoundES, total[0], ESINDMULServiceTotal1);
            //Service ES IND SERVICIO
            total = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(deliveryRoundES, ESINDServicioName);
            var actualTotal1 = total[0].Replace("Total : ", "").Trim();
            Assert.AreEqual(ESINDServiceTotal, actualTotal1, "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", ESINDServicioName, deliveryRoundES, total[0], ESINDServiceTotal);
            //Service ES MUL SERVICIO
            total = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(deliveryRoundES, ESMULServicioName);
            var actualTotal2 = total[0].Replace("Total : ", "").Trim();
            Assert.AreEqual(ESMULServiceTotal, actualTotal2, "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", ESMULServicioName, deliveryRoundES, total[0], ESMULServiceTotal);
            //DELIVERY WITH SERVICIO ES IND = 5, SERVICIO ES MUL = 10 & SERVICIO INDMUL = 10
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, ESDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDServicioName, quantity5);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESMULServicioName, quantity10);
            dispatchPage.AddQuantityOnPrevisonalQuantity(ESINDMULServicioName, quantity10);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isDispatchValidated = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isDispatchValidated, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESMULServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDMULServicioName);

            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundES);
            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service ES IND MUL SERVICIO
            total = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(deliveryRoundES, ESINDMULServicioName);
            actualTotal = total[0].Replace("Total : ", "").Trim();
            Assert.AreEqual(ESINDMULServiceTotal2, actualTotal, "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", ESINDMULServicioName, deliveryRoundES, total[0], ESINDMULServiceTotal2);
            //Service ES IND SERVICIO
            total = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(deliveryRoundES, ESINDServicioName);
            actualTotal1 = total[0].Replace("Total : ", "").Trim();
            Assert.AreEqual(ESINDServiceTotal, actualTotal1, "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", ESINDServicioName, deliveryRoundES, total[0], ESINDServiceTotal);
            //Service ES MUL SERVICIO
            total = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(deliveryRoundES, ESMULServicioName);
            actualTotal2 = total[0].Replace("Total : ", "").Trim();
            Assert.AreEqual(ESMULServiceTotal, actualTotal2, "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", ESMULServicioName, deliveryRoundES, total[0], ESMULServiceTotal);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckTotal_MethodPaxPerPackaging()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string TLSServiceTotal1 = "2700 g 300 g";
            string TLSServiceTotal2 = "3600 g";

            string TLSServiceQuantity1 = "10";
            string TLSServiceQuantity2 = "12";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //DELIVERY WITH SERVICIO TLS = 10
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, TLSServicioName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, TLSServiceQuantity1);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isDispatchValidated = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isDispatchValidated, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, TLSDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();

            //Service TLS SERVICIO
            var totalDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(TLSDeliveryRound, TLSServicioName);
            Assert.IsTrue(totalDisplayed[0].Contains(TLSServiceTotal1), "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", TLSServicioName, TLSDeliveryRound, totalDisplayed[0], TLSServiceTotal1);

            //DELIVERY WITH SERVICIO TLS = 12
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, TLSServicioName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, TLSServiceQuantity2);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            bool isDisptachValidatedByColorDay = previsionnalQty.IsValidatedByColorDay();
            Assert.IsTrue(isDisptachValidatedByColorDay, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, TLSDeliveryRound);
            setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();

            //Service TLS SERVICIO
            totalDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(TLSDeliveryRound, TLSServicioName);
            Assert.IsTrue(totalDisplayed[0].Contains(TLSServiceTotal2), "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", TLSServicioName, TLSDeliveryRound, totalDisplayed[0], TLSServiceTotal2);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckTotal_MethodPaxPerPackagingGrouped()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string TLSGRTotal = "1200 g 300 g";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundTLSGR);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service TLS GR 1 SERVICIO
            var totalDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(deliveryRoundTLSGR, "grouped by recipe");
            Assert.IsTrue(totalDisplayed[0].Contains(TLSGRTotal), "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", "grouped by recipe", deliveryRoundTLSGR, totalDisplayed[0], TLSGRTotal);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckTotal_MethodByRecipeVariant()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string pfServiceTotal = "1500 g 750 g 150 g";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRoundPF);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service PF SERVICIO
            var totalDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(deliveryRoundPF, servicioPFName);
            Assert.IsTrue(totalDisplayed[0].Contains(pfServiceTotal), "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", servicioPFName, deliveryRoundPF, totalDisplayed[0], pfServiceTotal);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckTotal_MethodByRecipeVariant2Recipes()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string CondrenServiceTotal1 = "1500 g 750 g";
            string CondrenServiceTotal2 = "200 g";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, condrenDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //Service Condren SERVICIO - RECETA CONDREN 1
            var totalDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(condrenDeliveryRound, condrenServicioName);
            Assert.IsTrue(totalDisplayed[1].Contains(CondrenServiceTotal1), "Dans Setup, onglet Delivery Round, le total du service {0}, associé à la rectte {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <Total : {4}>", condrenServicioName, condrenReceta1Name, condrenDeliveryRound, totalDisplayed[1], CondrenServiceTotal1);

            //Service Condren SERVICIO - RECETA CONDREN 2
            totalDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(condrenDeliveryRound, condrenServicioName);
            Assert.IsTrue(totalDisplayed[0].Contains(CondrenServiceTotal2), "Dans Setup, onglet Delivery Round, le total du service {0}, associé à la rectte {1}, dans le delivery round {2} ne correspond pas : <{3}> au lieu de <Total : {4}>", condrenServicioName, condrenReceta2Name, condrenDeliveryRound, totalDisplayed[0], CondrenServiceTotal2);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Display_CheckTotal_MethodRecipeVariantAndGroupedByRecipe()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string barentinBulkPackagingTotal = "6300 g 200 g";
            string barentinGroupedByTotal = "5000 g 1000 g";
            string barentinAdultoServicioTotal = "1500 g";
            string barentinCollegioServicioTotal = "4800 g 200 g";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, barentinDeliveryRound);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();

            //bulk packaging
            var totalDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(barentinDeliveryRound, "bulk packaging");
            Assert.IsTrue(totalDisplayed[0].Contains(barentinBulkPackagingTotal), "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", "bulk packaging", barentinDeliveryRound, totalDisplayed[0], barentinBulkPackagingTotal);

            //Grouped by
            totalDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(barentinDeliveryRound, "grouped by recipe");
            Assert.IsTrue(totalDisplayed[0].Contains(barentinGroupedByTotal), "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", "grouped by recipe", barentinDeliveryRound, totalDisplayed[0], barentinGroupedByTotal);

            //SERVICIO BARENTIN ADULTO 
            totalDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(barentinDeliveryRound, barentinServicioAdulto);
            Assert.IsTrue(totalDisplayed[0].Contains(barentinAdultoServicioTotal), "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", barentinServicioAdulto, barentinDeliveryRound, totalDisplayed[0], barentinAdultoServicioTotal);

            //SERVICIO BARENTIN COLLEGIO
            totalDisplayed = setupDeliveryRoundTabPage.GetDeliveryRoundServiceTotals(barentinDeliveryRound, barentinServicioCollegio);
            Assert.IsTrue(totalDisplayed[0].Contains(barentinCollegioServicioTotal), "Dans Setup, onglet Delivery Round, le total du service {0}, dans le delivery round {1} ne correspond pas : <{2}> au lieu de <Total : {3}>", barentinServicioCollegio, barentinDeliveryRound, totalDisplayed[0], barentinCollegioServicioTotal);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckNbLabel()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //prérequis //validation dispatch
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get nb label by recipe name, menu name and foodpack name


            Assert.AreEqual("4", setupPage.GetNbLabel(csvData, TLSRecetaName, "6", foodpackB6),
                "Sur le fichier csv, le nb Label ne correspond pas.");

            Assert.AreEqual("2", setupPage.GetNbLabel(csvData, recetaTLSGRName, "6", foodpackB6),
                "Sur le fichier csv, le nb Label ne correspond pas.");
            Assert.AreEqual("3", setupPage.GetNbLabel(csvData, recetaTLSGRName, "1", foodpackB1),
               "Sur le fichier csv, le nb Label ne correspond pas.");

            Assert.AreEqual("1", setupPage.GetNbLabel(csvData, recetaPFName, "10", foodpackBACGN),
                "Sur le fichier csv, le nb Label ne correspond pas.");
            Assert.AreEqual("1", setupPage.GetNbLabel(csvData, recetaPFName, "5", foodpackBACGNhalf),
                "Sur le fichier csv, le nb Label ne correspond pas.");
            Assert.AreEqual("(+1)", setupPage.GetNbLabel(csvData, recetaPFName, "", ""),
                "Sur le fichier csv, le nb Label ne correspond pas.");

            Assert.AreEqual("1", setupPage.GetNbLabel(csvData, condrenReceta1Name, "10", foodpackBACGN),
                "Sur le fichier csv, le nb Label ne correspond pas.");
            Assert.AreEqual("1", setupPage.GetNbLabel(csvData, condrenReceta1Name, "5", foodpackBACGNhalf),
                "Sur le fichier csv, le nb Label ne correspond pas.");
            Assert.AreEqual("2", setupPage.GetNbLabel(csvData, condrenReceta2Name, "1", foodpackB1),
                "Sur le fichier csv, le nb Label ne correspond pas.");

            Assert.AreEqual("1", setupPage.GetNbLabel(csvData, barentinReceta1, "10", foodpackBACGN),
                "Sur le fichier csv, le nb Label ne correspond pas.");
            Assert.AreEqual("3", setupPage.GetNbLabel(csvData, barentinReceta1, "16", foodpackBACGN),
                "Sur le fichier csv, le nb Label ne correspond pas.");
            Assert.AreEqual("(+2)", setupPage.GetNbLabel(csvData, barentinReceta1, "", ""),
                "Sur le fichier csv, le nb Label ne correspond pas.");
            Assert.AreEqual("1", setupPage.GetNbLabel(csvData, barentinRecetaNP2, "50", foodpackCRT),
                "Sur le fichier csv, le nb Label ne correspond pas.");
            Assert.AreEqual("1", setupPage.GetNbLabel(csvData, barentinRecetaNP2, "10", foodpackSEAU),
                "Sur le fichier csv, le nb Label ne correspond pas.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckSiteName()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var date = DateUtils.Now.AddYears(-1);
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, date);
            var setupPage = setupFilterAndFavoritesPage.Validate();


            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ClearDownloads();
            DeleteAllFileDownload();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get site name by recipe name, menu name and foodpack name
            var siteName1 = setupPage.GetSiteName(csvData, recetaESName, ESINDMenuName, foodpackIndividual);
            Assert.AreEqual(siteName1, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName2 = setupPage.GetSiteName(csvData, recetaESName, ESINDMULMenuName, foodpackMultiporcion);
            Assert.AreEqual(siteName2, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName3 = setupPage.GetSiteName(csvData, recetaESName, ESMULMenuName, foodpackMultiporcion);
            Assert.AreEqual(siteName3, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName4 = setupPage.GetSiteName(csvData, TLSRecetaName, TLSMenuName, foodpackB6);
            Assert.AreEqual(siteName4, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName5 = setupPage.GetSiteName(csvData, recetaTLSGRName, deliveryTLSGRName, foodpackB6);
            Assert.AreEqual(siteName5, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName6 = setupPage.GetSiteName(csvData, recetaTLSGRName, deliveryTLSGRName, foodpackB1);
            Assert.AreEqual(siteName6, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName7 = setupPage.GetSiteName(csvData, recetaPFName, menuPFName, foodpackBACGN);
            Assert.AreEqual(siteName7, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName8 = setupPage.GetSiteName(csvData, recetaPFName, menuPFName, foodpackBACGNhalf);
            Assert.AreEqual(siteName8, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName9 = setupPage.GetSiteName(csvData, recetaPFName, menuPFName, "");
            Assert.AreEqual(siteName9, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName10 = setupPage.GetSiteName(csvData, condrenReceta1Name, condrenMenuName, foodpackBACGN);
            Assert.AreEqual(siteName10, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName11 = setupPage.GetSiteName(csvData, condrenReceta1Name, condrenMenuName, foodpackBACGNhalf);
            Assert.AreEqual(siteName11, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName12 = setupPage.GetSiteName(csvData, condrenReceta2Name, condrenMenuName, foodpackB1);
            Assert.AreEqual(siteName12, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName13 = setupPage.GetSiteName(csvData, barentinReceta1, barentinMenuCollegio, foodpackBACGN);
            Assert.AreEqual(siteName13, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName14 = setupPage.GetSiteName(csvData, barentinReceta1, barentinMenuAdulto, foodpackBACGN);
            Assert.AreEqual(siteName14, siteACE, "Sur le fichier csv, le site name ne correspond pas.");

            var siteName15 = setupPage.GetSiteName(csvData, barentinReceta1, barentinMenuCollegio, "");
            Assert.AreEqual(siteName15, siteACE, "Sur le fichier csv, le site name ne correspond pas.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckMenuName()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //string deliveryTLSGRNames = "ES DELIVERY";
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);
            var resultat = setupPage.IsMenuNameInExcel(csvData, ESINDMenuName);
            //Get menu name

            Assert.IsTrue(setupPage.IsMenuNameInExcel(csvData, TLSMenuName), "Sur le fichier csv, le menu name n'est pas affiché.");
            Assert.IsTrue(setupPage.IsMenuNameInExcel(csvData, deliveryTLSGRName), "Sur le fichier csv, le menu name n'est pas affiché.");

            Assert.IsTrue(setupPage.IsMenuNameInExcel(csvData, menuPFName), "Sur le fichier csv, le menu name n'est pas affiché.");
            Assert.IsTrue(setupPage.IsMenuNameInExcel(csvData, menuPFName), "Sur le fichier csv, le menu name n'est pas affiché.");
            Assert.IsTrue(setupPage.IsMenuNameInExcel(csvData, menuPFName), "Sur le fichier csv, le menu name n'est pas affiché.");

            Assert.IsTrue(setupPage.IsMenuNameInExcel(csvData, condrenMenuName), "Sur le fichier csv, le menu name n'est pas affiché.");
            Assert.IsTrue(setupPage.IsMenuNameInExcel(csvData, condrenMenuName), "Sur le fichier csv, le menu name n'est pas affiché.");
            Assert.IsTrue(setupPage.IsMenuNameInExcel(csvData, condrenMenuName), "Sur le fichier csv, le menu name n'est pas affiché.");

            Assert.IsTrue(setupPage.IsMenuNameInExcel(csvData, barentinMenuAdulto), "Sur le fichier csv, le menu name n'est pas affiché.");
            Assert.IsTrue(setupPage.IsMenuNameInExcel(csvData, barentinMenuCollegio), "Sur le fichier csv, le menu name n'est pas affiché.");
            Assert.IsTrue(setupPage.IsMenuNameInExcel(csvData, barentinMenuCollegio), "Sur le fichier csv, le menu name n'est pas affiché.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckRecipeName()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var date = DateUtils.Now.AddYears(-1);
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, date);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();
            setupPage.ClearDownloads();
            DeleteAllFileDownload();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get recipe name by menu name
            var isRecipeNameInMenu = setupPage.GetRecipesNamesList(csvData, ESINDMenuName).Any(recipe => recetaESName.Contains(recipe));
            Assert.IsTrue(isRecipeNameInMenu, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            isRecipeNameInMenu = setupPage.GetRecipesNamesList(csvData, ESINDMULMenuName).Any(recipe => recetaESName.Contains(recipe));
            Assert.IsTrue(isRecipeNameInMenu, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            isRecipeNameInMenu = setupPage.GetRecipesNamesList(csvData, ESMULMenuName).Any(recipe => recetaESName.Contains(recipe));
            Assert.IsTrue(isRecipeNameInMenu, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            isRecipeNameInMenu = setupPage.GetRecipesNamesList(csvData, TLSMenuName).Any(recipe => TLSRecetaName.Contains(recipe));
            Assert.IsTrue(isRecipeNameInMenu, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            isRecipeNameInMenu = setupPage.GetRecipesNamesList(csvData, deliveryTLSGRName).Any(recipe => recetaTLSGRName.Contains(recipe));
            Assert.IsTrue(isRecipeNameInMenu, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            isRecipeNameInMenu = setupPage.GetRecipesNamesList(csvData, menuPFName).Any(recipe => recetaPFName.Contains(recipe));
            Assert.IsTrue(isRecipeNameInMenu, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            isRecipeNameInMenu = setupPage.GetRecipesNamesList(csvData, condrenMenuName).Any(recipe => condrenReceta1Name.Contains(recipe));
            Assert.IsTrue(isRecipeNameInMenu, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            isRecipeNameInMenu = setupPage.GetRecipesNamesList(csvData, condrenMenuName).Any(recipe => condrenReceta2Name.Contains(recipe));
            Assert.IsTrue(isRecipeNameInMenu, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            isRecipeNameInMenu = setupPage.GetRecipesNamesList(csvData, barentinMenuCollegio).Any(recipe => barentinReceta1.Contains(recipe));
            Assert.IsTrue(isRecipeNameInMenu, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            isRecipeNameInMenu = setupPage.GetRecipesNamesList(csvData, barentinMenuAdulto).Any(recipe => barentinReceta1.Contains(recipe));
            Assert.IsTrue(isRecipeNameInMenu, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckCommercialRecipeName()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var date = DateUtils.Now.AddYears(-1);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, date);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            //setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.GoToSetupDeliveryRoundTab();
            setupPage.ClearDownloads();
            DeleteAllFileDownload();
            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get recipe name by menu name
            var commercialRecipesNames = setupPage.GetCommercialRecipesNamesList(csvData, ESINDMenuName);
            var verifESINDMenuName = commercialRecipesNames.Any(recipe => recetaESName.Contains(recipe));
            Assert.IsTrue(verifESINDMenuName, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            var commercialRecipesNamesESINDMUL = setupPage.GetCommercialRecipesNamesList(csvData, ESINDMULMenuName);
            var verifESINDMULMenuName = commercialRecipesNamesESINDMUL.Any(recipe => recetaESName.Contains(recipe));
            Assert.IsTrue(verifESINDMULMenuName, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            var commercialRecipesNamesESMUL = setupPage.GetCommercialRecipesNamesList(csvData, ESMULMenuName);
            var verifESMULMenuName = commercialRecipesNamesESMUL.Any(recipe => recetaESName.Contains(recipe));
            Assert.IsTrue(verifESMULMenuName, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            var commercialRecipesNamesTLSGR = setupPage.GetCommercialRecipesNamesList(csvData, deliveryTLSGRName);
            var verifTLSGRName = commercialRecipesNamesTLSGR.Any(recipe => recetaTLSGRName.Contains(recipe));
            Assert.IsTrue(verifTLSGRName, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            var commercialRecipesNamesPF = setupPage.GetCommercialRecipesNamesList(csvData, menuPFName);
            var verifPFName = commercialRecipesNamesPF.Any(recipe => recetaPFName.Contains(recipe));
            Assert.IsTrue(verifPFName, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            var commercialRecipesNamesCondren = setupPage.GetCommercialRecipesNamesList(csvData, condrenMenuName);
            var verifCondrenReceta1Name = commercialRecipesNamesCondren.Any(recipe => condrenReceta1Name.Contains(recipe));
            Assert.IsTrue(verifCondrenReceta1Name, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            var verifCondrenReceta2Name = commercialRecipesNamesCondren.Any(recipe => condrenReceta2Name.Contains(recipe));
            Assert.IsTrue(verifCondrenReceta2Name, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            var commercialRecipesNamesBarentinCollegio = setupPage.GetCommercialRecipesNamesList(csvData, barentinMenuCollegio);
            var verifBarentinReceta1 = commercialRecipesNamesBarentinCollegio.Any(recipe => barentinReceta1.Contains(recipe));
            Assert.IsTrue(verifBarentinReceta1, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");

            var commercialRecipesNamesBarentinAdulto = setupPage.GetCommercialRecipesNamesList(csvData, barentinMenuAdulto);
            var verifBarentinReceta1Adulto = commercialRecipesNamesBarentinAdulto.Any(recipe => barentinReceta1.Contains(recipe));
            Assert.IsTrue(verifBarentinReceta1Adulto, "Sur le fichier csv, le recipe name ne correspond pas au menu name.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckGuestType()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();
            string menuVariantCollegioForACE = TestContext.Properties["Production_MenuVariant2ForACE"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get guest type name by menu name
            Assert.IsTrue(menuVariantCollegioForACE.Contains(setupPage.GetGuestType(csvData, barentinMenuCollegio)),
                "Sur le fichier csv, le Guest type ne correspond pas.");
            Assert.IsTrue(menuVariantForACE.Contains(setupPage.GetGuestType(csvData, barentinMenuAdulto)),
                "Sur le fichier csv, le Guest type ne correspond pas.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckMealType()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();
            string menuVariantCollegioForACE = TestContext.Properties["Production_MenuVariant2ForACE"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get guest type name by menu name
            Assert.IsTrue(menuVariantCollegioForACE.Contains(setupPage.GetMealType(csvData, barentinMenuCollegio)),
                "Sur le fichier csv, le Meal type ne correspond pas.");
            Assert.IsTrue(menuVariantForACE.Contains(setupPage.GetMealType(csvData, barentinMenuAdulto)),
                "Sur le fichier csv, le Meal type ne correspond pas.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckNbPAXPerPackaging()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get site name by recipe name, menu name and foodpack name
            Assert.IsTrue(setupPage.GetNbPAXPerPackaging(csvData, TLSMenuName, foodpackB6).Equals("6"),
                "Sur le fichier csv, le nb PAX / packaging name ne correspond pas.");
            Assert.IsTrue(setupPage.GetNbPAXPerPackaging(csvData, deliveryTLSGRName, foodpackB6).Equals("6"),
                "Sur le fichier csv, le nb PAX / packaging name ne correspond pas.");
            Assert.IsTrue(setupPage.GetNbPAXPerPackaging(csvData, deliveryTLSGRName, foodpackB1).Equals("1"),
                "Sur le fichier csv, le nb PAX / packaging name ne correspond pas.");

            Assert.IsTrue(setupPage.GetNbPAXPerPackaging(csvData, menuPFName, foodpackBACGN).Equals("10"),
                "Sur le fichier csv, le nb PAX / packaging name ne correspond pas.");
            Assert.IsTrue(setupPage.GetNbPAXPerPackaging(csvData, menuPFName, foodpackBACGNhalf).Equals("5"),
                "Sur le fichier csv, le nb PAX / packaging name ne correspond pas.");

            Assert.IsTrue(setupPage.GetNbPAXPerPackaging(csvData, condrenMenuName, foodpackBACGN).Equals("10"),
                "Sur le fichier csv, le nb PAX / packaging name ne correspond pas.");
            Assert.IsTrue(setupPage.GetNbPAXPerPackaging(csvData, condrenMenuName, foodpackBACGNhalf).Equals("5"),
                "Sur le fichier csv, le nb PAX / packaging name ne correspond pas.");
            Assert.IsTrue(setupPage.GetNbPAXPerPackaging(csvData, condrenMenuName, foodpackB1).Equals("1"),
                "Sur le fichier csv, le nb PAX / packaging name ne correspond pas.");

            Assert.IsTrue(setupPage.GetNbPAXPerPackaging(csvData, barentinMenuAdulto, foodpackBACGN).Equals("10"),
                "Sur le fichier csv, le nb PAX / packaging name ne correspond pas.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckFoodPackName()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var date = DateUtils.Now.AddYears(-1);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, date);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();
            setupPage.ClearDownloads();
            DeleteAllFileDownload();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get site name by recipe name, menu name and foodpack name
            Assert.IsTrue(foodpackIndividual == setupPage.GetFoodPackName(csvData, ESINDMenuName, "1"), "Sur le fichier csv, le foodpack ne correspond pas au menu.");
            Assert.IsTrue(foodpackMultiporcion == setupPage.GetFoodPackName(csvData, ESINDMULMenuName, "15"), "Sur le fichier csv, le foodpack ne correspond pas au menu.");
            Assert.IsTrue(foodpackMultiporcion == setupPage.GetFoodPackName(csvData, ESMULMenuName, "10"), "Sur le fichier csv, le foodpack ne correspond pas au menu.");

            Assert.IsTrue(foodpackB6 == setupPage.GetFoodPackName(csvData, TLSMenuName, portionsFoodpackB6), "Sur le fichier csv, le foodpack ne correspond pas au menu.");

            Assert.IsTrue(foodpackB6 == setupPage.GetFoodPackName(csvData, deliveryTLSGRName, portionsFoodpackB6), "Sur le fichier csv, le foodpack ne correspond pas au menu.");
            Assert.IsTrue(foodpackB1 == setupPage.GetFoodPackName(csvData, deliveryTLSGRName, portionsFoodpackB1), "Sur le fichier csv, le foodpack ne correspond pas au menu.");

            Assert.IsTrue(foodpackBACGN == setupPage.GetFoodPackName(csvData, menuPFName, portionsFoodpackBACGN), "Sur le fichier csv, le foodpack ne correspond pas au menu.");
            Assert.IsTrue(foodpackBACGNhalf == setupPage.GetFoodPackName(csvData, menuPFName, portionsFoodpackBACGNhalf), "Sur le fichier csv, le foodpack ne correspond pas au menu.");

            Assert.IsTrue(foodpackBACGN == setupPage.GetFoodPackName(csvData, condrenMenuName, portionsFoodpackBACGN), "Sur le fichier csv, le foodpack ne correspond pas au menu.");
            Assert.IsTrue(foodpackBACGNhalf == setupPage.GetFoodPackName(csvData, condrenMenuName, portionsFoodpackBACGNhalf), "Sur le fichier csv, le foodpack ne correspond pas au menu.");
            Assert.IsTrue(foodpackB1 == setupPage.GetFoodPackName(csvData, condrenMenuName, portionsFoodpackB1), "Sur le fichier csv, le foodpack ne correspond pas au menu.");

            Assert.IsTrue(foodpackBACGN == setupPage.GetFoodPackName(csvData, barentinMenuCollegio, "16"), "Sur le fichier csv, le foodpack ne correspond pas au menu.");
            Assert.IsTrue(foodpackBACGN == setupPage.GetFoodPackName(csvData, barentinMenuAdulto, portionsFoodpackBACGN), "Sur le fichier csv, le foodpack ne correspond pas au menu.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckSiteAdress()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteACEAddress = "LG Playa Honda";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act

            //1. Check if adress on site
            var settingsSitesPage = homePage.GoToParameters_Sites();
            settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, siteACE);
            settingsSitesPage.ClickOnFirstSite();
            settingsSitesPage.ClickToInformations();
            var siteAddress = settingsSitesPage.GetAddress();
            if (!siteAddress.Equals(siteACEAddress))
            {
                settingsSitesPage.SetAddress(siteACEAddress);
            }

            //2. Export setup file to check site adress
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get site name by recipe name, menu name and foodpack name
            Assert.IsFalse(setupPage.GetSiteAddress(csvData, siteACE).Any(address => address != siteACEAddress), "Sur le fichier csv, le 'site address' ne correspond pas au site.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckSanitaryAgreement()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteACESanitaryAgreement = "26.02075/GC";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act

            //1. Check if sanitary agreement on site
            var settingsSitesPage = homePage.GoToParameters_Sites();
            settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, siteACE);
            settingsSitesPage.ClickOnFirstSite();
            settingsSitesPage.ClickToInformations();

            if (!settingsSitesPage.GetSanitaryAgreement().Equals(siteACESanitaryAgreement))
            {
                settingsSitesPage.SetSanitaryAgreement(siteACESanitaryAgreement);
            }

            //2. Export setup file to check site sanitary agreement
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get site name by recipe name, menu name and foodpack name
            Assert.IsFalse(setupPage.GetSiteSanitaryAgreement(csvData, siteACE).Any(agreement => agreement != siteACESanitaryAgreement), "Sur le fichier csv, le le sanitary agreement ne correspond pas au site.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckRecipeComment()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string ESrecipeComment = "ES RECETA COMMENT";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Recipe Comment";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act

            //1. Add a comment in recipe
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaESName);
            var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeGeneralInfosPage.ClickOnEditInformation();

            if (!recipeGeneralInfosPage.GetRecipeInformationComment().Equals(ESrecipeComment))
            {
                recipeGeneralInfosPage.SetRecipeInformationComment(ESrecipeComment);
            }

            //2. Export setup file to check recipe comment
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get site name by recipe name, menu name and foodpack name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, recetaESName).Any(agreement => agreement != ESrecipeComment), "Sur le fichier csv, le 'recipe comment' ne correspond pas au commentaire ajouté dans la recette {0}.", recetaESName);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckNutricPerPortion()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string TLSEnergyKCal = "600";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Nutric per portion";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act

            //1. Add Energy KCAL in recipe
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, TLSRecetaName);
            var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            recipeVariantPage.ClickOnDieteticTab();
            if (!recipeVariantPage.GetEnergyKCal().Equals(TLSEnergyKCal))
            {
                recipeVariantPage.SetEnergyKCal(TLSEnergyKCal);
            }

            //2. Export setup file to check recipe comment
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get site name by recipe name, menu name and foodpack name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, TLSRecetaName).Any(energy => energy != TLSEnergyKCal), "Sur le fichier csv, le 'Nutric Per Portion' ne correspond pas à l'EnergyKCal dans l'onglet 'Dietetic' de la recette {0}.", TLSRecetaName);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckNutricPer100g()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string TLSEnergyKCal = "600";
            string TLSEnergyKCalPer100g = "400"; //600*100/150 (portion de 150 g)
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Nutric per 100G";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act

            //1. Add Energy KCAL in recipe
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, TLSRecetaName);
            var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            recipeVariantPage.ClickOnDieteticTab();
            if (!recipeVariantPage.GetEnergyKCal().Equals(TLSEnergyKCal))
            {
                recipeVariantPage.SetEnergyKCal(TLSEnergyKCal);
            }

            //2. Export setup file to check recipe comment
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get site name by recipe name, menu name and foodpack name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, TLSRecetaName).Any(energy => energy != TLSEnergyKCalPer100g), "Sur le fichier csv, le 'Nutric Per 100g' ne correspond pas à l'EnergyKCal dans l'onglet 'Dietetic' de la recette {0} et/ou de sa portion.", TLSRecetaName);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckComposition()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Composition";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get site name by recipe name, menu name and foodpack name
            Assert.IsTrue(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, recetaTLSGRName).All(composition => composition.Contains(itemTest)), "Sur le fichier csv, la 'Composition' ne correspond pas à la composition de la recette {0}.", recetaTLSGRName);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckIntolerance()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string allergen = TestContext.Properties["RecipeAllergen"].ToString();
            string columnToTestName = "Intolerence";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            // 1. Add allergen in recipe item (itemTest created in test PR_SETUP_Create_Items)
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemTest);
            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            var itemIntolerancePage = itemGeneralInformationPage.ClickOnIntolerancePage();
            itemIntolerancePage.AddAllergen(allergen);

            //2. Check allergen present in recipe intolerance tab (recipe PF created in test PR_SETUP_Create_Datas_MethodByRecipeVariant)
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaPFName);
            var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            var recipeIntolerancePage = recipeVariantPage.ClickOnIntoleranceTab();
            Assert.IsTrue(recipeIntolerancePage.IsRecipeAllergenChecked(allergen), "L'allergène {0} selectionné sur l'item {1} ne remonte pas bien sur la recette {2}", allergen, itemTest, recetaPFName);

            //3. Check in export file allergen
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get site name by recipe name, menu name and foodpack name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, recetaPFName).Any(intolerence => intolerence != allergen),
                "Sur le fichier csv, l'Intolerence ne correspond pas à l'allergen {0} sélectionné dans l'item {1} et présent dans la recette {2}.", allergen, itemTest, recetaPFName);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckProductionDate()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Production day";
            var today = DateUtils.Now.ToString("yyyy-MM-dd");

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Check in export file production date
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get production days
            Assert.IsFalse(setupPage.GetExportResults(csvData, columnToTestName).Any(productionDay => productionDay != today),
                "Sur le fichier csv, la 'production date' ne correspond pas à la date selectionnée lors de l'export setup, date attendue : {0}", today);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckMenuDay()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Menu day";
            var date = DateUtils.Now.AddDays(-1);
            var yesterday = date.ToString("yyyy-MM-dd");
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Check in export file menu day
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.ClearDownloads();

            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.StartDate, date);


            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get production days
            Assert.IsFalse(setupPage.GetExportResults(csvData, columnToTestName).Any(menuDay => menuDay != yesterday),
                "Sur le fichier csv, la 'menu day' ne correspond pas à la date selectionnée lors de la saisie à l'initialisation de la page setup (filtre), date attendue : {0}", yesterday);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckExpiryDay()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Expiry day";
            var numberOfDays = "2";
            string expiryDate = DateUtils.Now.AddDays(+2).ToString("dd/MM/yyyy");

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            // 1. Définir la recipe expiry date
            homePage.SetProductionSettingsRecipeExpiryDateValue(numberOfDays);

            // 2. Check in export file production date
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get production days
            Assert.IsTrue(setupPage.GetExportResults(csvData, columnToTestName).Any(expiryDay => expiryDay != expiryDate),
                "Sur le fichier csv, le 'expiry day' ne correspond pas à la date de production additionnée au nombre de jours saisis dans le global settings, date attendue : {0}", expiryDate);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckItemCaracteristic()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string caracteristic = "caracteristic";
            string columnToTestName = "Item caracteristics";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            // 1. Add caracteristic in item (itemTest created in test PR_SETUP_Create_Items)
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemTest);
            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            var itemLabelPage = itemGeneralInformationPage.ClickOnLabelPage();
            itemLabelPage.EditLabel(caracteristic);
            itemLabelPage.Validate();

            //2. Check in export file caracteristic
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            //SetupDeliveryRoundTabPage setupDeliveryRoundPage = setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get caracteristic by recipe name
            List<string> result = setupPage.GetExportResultsByRecipe(csvData, columnToTestName, recetaPFName);
            foreach (string r in result)
            {
                Assert.AreEqual(r, caracteristic, "Sur le fichier csv, l'Item carcteristic ne correspond pas au caracteristic {0} dans l'item {1}, présent dans la recette {2}.", caracteristic, itemTest, recetaPFName);
            }
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckRecipeConserveMode()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Conserve mode";
            string recipeConserveMode = "refrigeration";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            homePage.ClearDownloads();

            //Act

            //1. Add conserve mode in recipe (edit information tab)
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, TLSRecetaName);
            var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeGeneralInfosPage.ClickOnEditInformation();
            if (!recipeGeneralInfosPage.GetConserveMode().Equals(recipeConserveMode))
            {
                recipeGeneralInfosPage.SetConserveMode(recipeConserveMode);
            }

            //2. Export setup file to check recipe comment
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get recipe cooking mode by recipe name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, TLSRecetaName).Any(conserveMode => conserveMode != recipeConserveMode), "Sur le fichier csv, le 'Conserve mode' ne correspond pas Conserve mode dans l'onglet 'Edit information' de la recette {0}.", TLSRecetaName);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckRecipeConserveTemperature()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Conserve temperature";
            string recipeConserveTemperature = "4°C";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act

            //1. Add conserve temperature in recipe (edit information tab)
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, TLSRecetaName);
            var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeGeneralInfosPage.ClickOnEditInformation();
            recipeGeneralInfosPage.SetConserveTemperature(recipeConserveTemperature);

            //2. Export setup file to check recipe comment
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get conserve temperature by recipe name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, TLSRecetaName).Any(conserveTp => conserveTp != recipeConserveTemperature), "Sur le fichier csv, la 'Conserve temperature' ne correspond pas Conserve temperature dans l'onglet 'Edit information' de la recette {0}.", TLSRecetaName);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckWarmingMethod()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Warming method";
            string warmingMode = "Four"; //120°C

            // Arrange
            var homePage = LogInAsAdmin();

            //Act

            //1. Add warming method in setting/production/food pack
            var parametersProduction = homePage.GoToParameters_ProductionPage();
            parametersProduction.GoToTab_Foodpack();

            if (!parametersProduction.CheckFoodPackExists(foodpackB1))
            {
                parametersProduction.CreateNewFoodPack(foodpackB1);
            }

            if (!parametersProduction.GetWarmingMode(foodpackB1).Equals(warmingMode))
            {
                parametersProduction.AddWarmingModeToFoodPack(foodpackB1, warmingMode);
            }

            //2. Export setup file to check recipe comment
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get warming method by recipe name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, condrenReceta2Name).Any(warmingMt => warmingMt != warmingMode), "Sur le fichier csv, la 'Warming method' ne correspond pas au Warming Mode ajouté dans le foodpack {0}.", foodpackB1);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckWarmingTemperature()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Warming temperature";
            string warmingTemperature = "120°C";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act

            //1. Add warming temperature in setting/production/food pack
            var parametersProduction = homePage.GoToParameters_ProductionPage();
            parametersProduction.GoToTab_Foodpack();

            if (!parametersProduction.CheckFoodPackExists(foodpackB1))
            {
                parametersProduction.CreateNewFoodPack(foodpackB1);
            }

            if (!parametersProduction.GetWarmingTemperature(foodpackB1).Equals(warmingTemperature))
            {
                parametersProduction.AddWarmingTemperatureToFoodPack(foodpackB1, warmingTemperature);
            }

            //2. Export setup file to check recipe comment
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get warming temperature by recipe name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, condrenReceta2Name).Any(warmingTp => warmingTp != warmingTemperature), "Sur le fichier csv, la 'Warming temperature' ne correspond pas au Warming temperature ajouté dans le foodpack {0}.", foodpackB1);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckRecipeDetailComment()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Comments";
            string recipeComment = "recipe detail comment";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act

            //1. Add recipe detail comment in recipe variant detail
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recetaPFName);
            var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

            if (!recipeVariantPage.GetComment(false).Equals(recipeComment))
            {
                recipeVariantPage.CloseCommentPopup();
                recipeVariantPage.ClickOnItemTab();
                recipeVariantPage.AddCommentToIngredient(recipeComment, false);
            }

            //2. Export setup file to check recipe comment
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get comment by recipe name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, recetaPFName).Any(comment => comment != recipeComment), "Sur le fichier csv, le Comment ne correspond pas au commentaire dans le détail de la recette {0}.", recetaPFName);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckDeliveryLabelComment()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Label comment";
            string deliveryLabelComment = "delivery label comment";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act

            //1. Add comment in delivery label comment
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, condrenDeliveryName);
            var deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
            var deliveryGeneralInformationPage = deliveryLoadingPage.ClickOnGeneralInformation();

            if (!deliveryGeneralInformationPage.GetLabelComment().Equals(deliveryLabelComment))
            {
                deliveryGeneralInformationPage.SetLabelComment(deliveryLabelComment);
            }

            //2. Export setup file to check recipe comment
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get label comment by recipe name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, condrenReceta1Name).Any(comment => comment != deliveryLabelComment), "Sur le fichier csv, le 'Label comment' ne correspond pas au commentaire ajouté dans la case 'Label comment' du delivery {0}.", condrenDeliveryName);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckDeliveryNoteComment()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Delivery note comment";
            string deliveryNoteComment = "delivery note comment";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act

            //1. Add comment in Specific comment printed on delivery notes
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, ESDeliveryName);
            var deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
            var deliveryGeneralInformationPage = deliveryLoadingPage.ClickOnGeneralInformation();

            if (!deliveryGeneralInformationPage.GetDeliveryNotesComment().Equals(deliveryNoteComment))
            {
                deliveryGeneralInformationPage.SetDeliveryNotesComment(deliveryNoteComment);
            }

            //2. Export setup file to check recipe comment
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get label comment by recipe name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, recetaESName).Any(comment => comment != deliveryNoteComment), "Sur le fichier csv, le 'Delivery note comment' ne correspond pas au commentaire ajouté dans la case 'Specific comment printed on delivery notes' du delivery {0}.", ESDeliveryName);
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Export_CheckCustomerDNComment()
        {
            //Prepare
            bool newVersionPrint = true;
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string columnToTestName = "Customer DN comment";
            string deliveryNoteComment = "customer delivery note comment";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act

            //1. Add comment in Specific comment printed on delivery notes
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, barentinCustomerName);
            var customerGeneralInformationPage = customersPage.SelectFirstCustomer();

            if (!customerGeneralInformationPage.GetCustomerDeliveryNoteComment().Equals(deliveryNoteComment))
            {
                customerGeneralInformationPage.SetCustomerDeliveryNoteComment(deliveryNoteComment);
            }

            //2. Export setup file to check recipe comment
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();

            setupPage.ClearDownloads();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);

            setupPage.GoToSetupDeliveryRoundTab();

            setupPage.ExportCSV(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = setupPage.GetCsvFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier CSV à ouvrir
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Get csv data (all lines)
            var csvData = setupPage.GetCSVData(filePath);

            //Get label comment by recipe name
            Assert.IsFalse(setupPage.GetExportResultsByRecipe(csvData, columnToTestName, barentinReceta1).Any(comment => comment != deliveryNoteComment), "Sur le fichier csv, le 'Customer DN comment' ne correspond pas au commentaire ajouté dans la case 'Delivery note comment' du customer {0}.", barentinCustomerName);
        }


        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Logo_DeliveryNote()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string directory = TestContext.Properties["DownloadsPath"].ToString();
            string site = "ACE";
            HomePage homePage = LogInAsAdmin();
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ClearDownloads();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, site);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, DateTime.Now);
            setupFilterAndFavoritesPage.GetDeliveryRound();
            //Arrange
            homePage.Navigate();
            ParametersSites SitesParams = homePage.GoToParameters_Sites();
            SitesParams.Filter(ParametersSites.FilterType.SearchSite, site);
            SitesParams.ClickOnFirstSite();
            //Configurer un logo sur un site dans les globals settings
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            SitesParams.UploadLogo(fiUpload);
            //return au setup delivery round page
            homePage.Navigate();
            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, site);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, DateTime.Now);
            setupFilterAndFavoritesPage.GetDeliveryRound();
            // purger le dossier Download
            foreach (string f in Directory.GetFiles(directory)) File.Delete(f);
            //print delivery note by recipes
            setupFilterAndFavoritesPage.GoPrintDeliveryNote();
            PrintReportPage reportPage = setupFilterAndFavoritesPage.PrintArchive(setupFilterAndFavoritesPage);
            setupFilterAndFavoritesPage.ClearDownloadFolder(directory);
            foreach (string f in Directory.GetFiles(directory)) if (f.EndsWith(".zip"))
                {
                    ZipFile.ExtractToDirectory(f, downloadsPath);
                    break;
                }
            List<FileInfo> files = new List<FileInfo>();
            foreach (string f in Directory.GetFiles(directory))
            {
                if (f.EndsWith(".zip") || f.EndsWith(".png")) continue;
                files.Add(new FileInfo(f));
            }
            for (int i = 0; i < files.Count; i++)
            {
                FileInfo fi = files[i];
                string currentFile = downloadsPath + "\\pizza_test.png";
                //3.Vérifier le logo sur le PDF
                PdfDocument document = PdfDocument.Open(fi.FullName);
                Page page1 = document.GetPage(1);
                List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
                Assert.AreEqual(1, images.Count, "Il n'existe aucune image.");
                IPdfImage image = images[0];
                File.WriteAllBytes(currentFile, image.RawBytes.ToArray<byte>());
                FileInfo fiTest;
                fiTest = new FileInfo(TestContext.TestDeploymentDir + "\\pizza_petite.png");
                Assert.IsTrue(fiTest.Exists, "fichier non trouvé (cas 2)");
                // comparaison fichiers
                MD5 md5 = MD5.Create();
                FileStream fsTest = File.Open(fiTest.FullName, FileMode.Open);
                FileStream fsPdf = File.Open(currentFile, FileMode.Open);
                Assert.AreEqual(System.Convert.ToBase64String(md5.ComputeHash(fsTest)), System.Convert.ToBase64String(md5.ComputeHash(fsPdf)), "fichiers différents");
                fsTest.Close();
                fsPdf.Close();
            }
        }


        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Logo_DeliveryNoteValorized()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Print SetUp_-_";

            var site = "ACE";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ClearDownloads();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, site);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, DateTime.Now);
            setupFilterAndFavoritesPage.GetDeliveryRound();

            //Arrange
            homePage.Navigate();
            ParametersSites SitesParams = homePage.GoToParameters_Sites();
            SitesParams.Filter(ParametersSites.FilterType.SearchSite, site);
            SitesParams.ClickOnFirstSite();

            //Configurer un logo sur un site dans les globals settings
            Console.WriteLine(TestContext.TestDeploymentDir + "\\pizza.png");
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            SitesParams.UploadLogo(fiUpload);

            //return au setup delivery round page
            homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, site);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, DateTime.Now);
            setupFilterAndFavoritesPage.GetDeliveryRound();

            //print delivery note valorized
            setupFilterAndFavoritesPage.GoPrintDeliveryValorized();

            PrintReportPage reportPage = setupFilterAndFavoritesPage.PrintItemGenerique(setupFilterAndFavoritesPage);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            //3.Vérifier le logo sur le PDF
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
            Assert.AreEqual(1, images.Count, "Il n'existe aucune image.");
            IPdfImage image = images[0];

            FileInfo fiPdf = new FileInfo(downloadsPath + "\\pizza_test.png");
            if (fiPdf.Exists)
            {
                fiPdf.Delete();
            }
            File.WriteAllBytes(downloadsPath + "\\pizza_test.png", image.RawBytes.ToArray<byte>());
            fiPdf.Refresh();
            Assert.IsTrue(fiPdf.Exists, "fichier non trouvé");

            FileInfo fiTest;
            fiTest = new FileInfo(TestContext.TestDeploymentDir + "\\pizza_petite.png");
            Assert.IsTrue(fiTest.Exists, "fichier non trouvé (cas 2)");

            // comparaison fichiers
            var md5 = MD5.Create();
            var fsTest = File.Open(fiTest.FullName, FileMode.Open);
            var fsPdf = File.Open(fiPdf.FullName, FileMode.Open);
            Assert.AreEqual(System.Convert.ToBase64String(md5.ComputeHash(fsTest)), System.Convert.ToBase64String(md5.ComputeHash(fsPdf)), "fichiers différents");

        }


        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Logo_DeliveryNotpteByServices()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Print SetUp_-_";
            string site = "ACE";
            HomePage homePage = LogInAsAdmin();
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ClearDownloads();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, site);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, DateTime.Now);
            setupFilterAndFavoritesPage.GetDeliveryRound();
            //Arrange
            homePage.Navigate();
            ParametersSites SitesParams = homePage.GoToParameters_Sites();
            SitesParams.Filter(ParametersSites.FilterType.SearchSite, site);
            SitesParams.ClickOnFirstSite();
            //Configurer un logo sur un site dans les globals settings
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            SitesParams.UploadLogo(fiUpload);
            //return au setup delivery round page
            homePage.Navigate();
            setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, site);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, DateTime.Now);
            setupFilterAndFavoritesPage.GetDeliveryRound();
            //print delivery Note by services
            setupFilterAndFavoritesPage.GoPrintDeliveryNoteByService();
            PrintReportPage reportPage = setupFilterAndFavoritesPage.PrintItemGenerique(setupFilterAndFavoritesPage);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string found = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fi = new FileInfo(found);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, $"{found} non généré");
            //3.Vérifier le logo sur le PDF
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
            Assert.AreEqual(1, images.Count, "Il n'existe aucune image.");
            IPdfImage image = images[0];
            FileInfo fiPdf = new FileInfo(downloadsPath + "\\pizza_test.png");
            if (fiPdf.Exists) fiPdf.Delete();
            File.WriteAllBytes(downloadsPath + "\\pizza_test.png", image.RawBytes.ToArray<byte>());
            fiPdf.Refresh();
            Assert.IsTrue(fiPdf.Exists, "fichier non trouvé");
            FileInfo fiTest;
            fiTest = new FileInfo(TestContext.TestDeploymentDir + "\\pizza_petite.png");
            Assert.IsTrue(fiTest.Exists, "fichier non trouvé (cas 2)");
            // comparaison fichiers
            MD5 md5 = MD5.Create();
            FileStream fsTest = File.Open(fiTest.FullName, FileMode.Open);
            FileStream fsPdf = File.Open(fiPdf.FullName, FileMode.Open);
            Assert.AreEqual(System.Convert.ToBase64String(md5.ComputeHash(fsTest)), System.Convert.ToBase64String(md5.ComputeHash(fsPdf)), "fichiers différents");
        }


        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Logo_DeliveryNoteByRecipes()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Print SetUp_-_";

            var site = "ACE";
            var homePage = LogInAsAdmin();
            ParametersSites SitesParams = homePage.GoToParameters_Sites();
            SitesParams.Filter(ParametersSites.FilterType.SearchSite, site);
            SitesParams.ClickOnFirstSite();

            //Configurer un logo sur un site dans les globals settings
            Console.WriteLine(TestContext.TestDeploymentDir + "\\pizza.png");
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            SitesParams.UploadLogo(fiUpload);

            //return au setup delivery round page
            homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, site);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, DateTime.Now);
            // valide le favory et va dans l'index setup
            setupFilterAndFavoritesPage.GetDeliveryRound();

            //print delivery Note by recipes
            setupFilterAndFavoritesPage.GoPrintDeliveryNoteByRecipes();

            PrintReportPage reportPage = setupFilterAndFavoritesPage.PrintItemGenerique(setupFilterAndFavoritesPage);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            //3.Vérifier le logo sur le PDF
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
            Assert.AreEqual(1, images.Count, "Il n'existe aucune image.");
            IPdfImage image = images[0];

            FileInfo fiPdf = new FileInfo(downloadsPath + "\\pizza_test.png");
            if (fiPdf.Exists)
            {
                fiPdf.Delete();
            }
            File.WriteAllBytes(downloadsPath + "\\pizza_test.png", image.RawBytes.ToArray<byte>());
            fiPdf.Refresh();
            Assert.IsTrue(fiPdf.Exists, "fichier non trouvé");

            FileInfo fiTest;
            fiTest = new FileInfo(TestContext.TestDeploymentDir + "\\pizza_petite.png");
            Assert.IsTrue(fiTest.Exists, "fichier non trouvé (cas 2)");

            // comparaison fichiers
            var md5 = MD5.Create();
            var fsTest = File.Open(fiTest.FullName, FileMode.Open);
            var fsPdf = File.Open(fiPdf.FullName, FileMode.Open);
            Assert.AreEqual(System.Convert.ToBase64String(md5.ComputeHash(fsTest)), System.Convert.ToBase64String(md5.ComputeHash(fsPdf)), "fichiers différents");
        }


        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Logo_FoodPack()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Print SetUp_-_";

            var site = "ACE";

            HomePage homePage = LogInAsAdmin();

            ParametersSites SitesParams = homePage.GoToParameters_Sites();
            SitesParams.Filter(ParametersSites.FilterType.SearchSite, site);
            SitesParams.ClickOnFirstSite();

            //Configurer un logo sur un site dans les globals settings
            Console.WriteLine(TestContext.TestDeploymentDir + "\\pizza.png");
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            SitesParams.UploadLogo(fiUpload);

            //return au setup delivery round page
            homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            DeleteAllFileDownload();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, site);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, DateTime.Now);
            // valide le favory et va dans l'index setup
            setupFilterAndFavoritesPage.GetDeliveryRound();
            setupFilterAndFavoritesPage.ClearDownloads();

            //print Food Pack report
            setupFilterAndFavoritesPage.GoPrintFoodPackReport();

            PrintReportPage reportPage = setupFilterAndFavoritesPage.PrintItemGenerique(setupFilterAndFavoritesPage);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            //3.Vérifier le logo sur le PDF
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
            Assert.AreEqual(1, images.Count, "Il n'existe aucune image.");
        }


        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Logo_DeliveryRound()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Print SetUp_-_";
            string site = "ACE";
            HomePage homePage = LogInAsAdmin();
            ParametersSites SitesParams = homePage.GoToParameters_Sites();
            SitesParams.Filter(ParametersSites.FilterType.SearchSite, site);
            SitesParams.ClickOnFirstSite();
            //Configurer un logo sur un site dans les globals settings
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            SitesParams.UploadLogo(fiUpload);
            //return au setup delivery round page
            homePage.Navigate();
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            DeleteAllFileDownload();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, site);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, DateTime.Now);
            // valide le favory et va dans l'index setup
            setupFilterAndFavoritesPage.GetDeliveryRound();
            setupFilterAndFavoritesPage.ClearDownloads();
            //print Delivery Round report
            setupFilterAndFavoritesPage.GoPrintDeliveryRound();
            PrintReportPage reportPage = setupFilterAndFavoritesPage.PrintItemGenerique(setupFilterAndFavoritesPage);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string found = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fi = new FileInfo(found);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, $"{found} non généré");
            //3.Vérifier le logo sur le PDF
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
            Assert.AreEqual(1, images.Count, "Il n'existe aucune image.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Logo_FoodPackDelivery()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Print SetUp_-_";
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            //arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            DeleteAllFileDownload();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.Validate();
            setupFilterAndFavoritesPage.ClearDownloads();
            setupFilterAndFavoritesPage.GoPrintFoodPackReport();
            PrintReportPage reportPage = setupFilterAndFavoritesPage.PrintItemGenerique(setupFilterAndFavoritesPage);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            //3.Vérifier le logo sur le PDF
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
            Assert.AreEqual(1, images.Count, "Il n'existe aucune image.");
        }


        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Logo_DeliveryRoundByRecipes()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Print SetUp_-_";

            var site = "ACE";
            var homePage = LogInAsAdmin();
            ParametersSites SitesParams = homePage.GoToParameters_Sites();
            SitesParams.Filter(ParametersSites.FilterType.SearchSite, site);
            SitesParams.ClickOnFirstSite();

            //Configurer un logo sur un site dans les globals settings
            Console.WriteLine(TestContext.TestDeploymentDir + "\\pizza.png");
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            SitesParams.UploadLogo(fiUpload);

            //return au setup delivery round page
            homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, site);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, DateTime.Now);
            // valide le favory et va dans l'index setup
            setupFilterAndFavoritesPage.GetDeliveryRound();
            setupFilterAndFavoritesPage.ClearDownloads();
            //print Delivery Round report
            setupFilterAndFavoritesPage.GoPrintDeliveryRoundByRecipes();

            PrintReportPage reportPage = setupFilterAndFavoritesPage.PrintItemGenerique(setupFilterAndFavoritesPage);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            //3.Vérifier le logo sur le PDF
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
            Assert.AreEqual(1, images.Count, "Il n'existe aucune image.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_DeliveriesFilter_ApresPlusieursUtilisations()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            // Filter for site ALC
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, "ALC");
            var selectedDeliveriesALC = setupFilterAndFavoritesPage.GetDeliveries();
            Assert.IsNotNull(selectedDeliveriesALC, "Le filtre 'Deliveries' doit retourner des résultats valides pour le site ALC après application des filtres.");
            Assert.IsTrue(selectedDeliveriesALC.Any(), "Il doit y avoir des livraisons sélectionnées pour le site ALC.");
            // Filter for site ACE
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, "ACE");
            var selectedDeliveriesACE = setupFilterAndFavoritesPage.GetDeliveries();
            Assert.IsNotNull(selectedDeliveriesACE, "Le filtre 'Deliveries' doit retourner des résultats valides pour le site ACE après application des filtres.");
            Assert.IsTrue(selectedDeliveriesACE.Any(), "Il doit y avoir des livraisons sélectionnées pour le site ACE.");
            // Filter for site CPU and expect no deliveries
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, "CPU");
            var buttonText = setupFilterAndFavoritesPage.GetDeliveriesDropdownState();
            Assert.AreEqual("Nothing Selected", buttonText, "Le bouton de livraison doit afficher 'Nothing Selected' pour le site CPU.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_RetourLigneNomDeliveryRound()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string deliveryRName = $"CONDREN INFORMATION {DateTime.Now.Day}{DateTime.Now.Month}{DateTime.Now.Hour}{DateTime.Now.Minute}{DateTime.Now.Second}";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Print SetUp_-_";
            DateTime startDate = DateTime.Now.AddDays(-2);
            DateTime endDate = DateTime.Now.AddDays(+31);
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            // delivery round  
            DeliveryRoundPage deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRName);
            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                // Create delivery round  
                DeliveryRoundCreateModalpage deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRName, siteACE, startDate, endDate);
                DeliveryRoundDeliveryPage deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(delivery2Name);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.ShowAll, true);
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRName);
            }
            else
            {
                DeliveryRoundDeliveryPage deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }
            //Assert
            string firstDeliveryRound = deliveryRoundPage.GetFirstDeliveryRound();
            Assert.AreEqual(deliveryRName, firstDeliveryRound, "Le delivery round n'a pas été créé.");
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            SetupPage setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.StartDate, DateUtils.Now.AddDays(-2));
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRName);
            setupFilterAndFavoritesPage.SelectDeliveryRound();
            setupFilterAndFavoritesPage.ClearDownloads();
            DeleteAllFileDownload();
            //print delivery Note by recipes
            setupFilterAndFavoritesPage.GoPrintDeliveryNoteByRecipes();
            PrintReportPage reportPage = setupFilterAndFavoritesPage.PrintItemGenerique2(setupFilterAndFavoritesPage);
            reportPage.Close();
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string found = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(found);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, $"{found} non généré");
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            List<string> wordStrings = words.Select(word => word.Text).ToList();
            string[] result = deliveryRName.Split(' ');
            bool IsRetourLigneEffectuer = setupFilterAndFavoritesPage.RetourLigneEffectuer(wordStrings, result);
            Assert.IsTrue(IsRetourLigneEffectuer, "il faudrait espacer suffisamment pour que les lignes delivery note by recipes se chevauchent pas.");
            setupPage.Filter(SetupPage.FilterType.Sites, siteACE);
            setupPage.Filter(SetupPage.FilterType.StartDate, DateUtils.Now.AddDays(-2));
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRName);
            setupFilterAndFavoritesPage.SelectDeliveryRound();
            setupFilterAndFavoritesPage.ClearDownloads();
            DeleteAllFileDownload();
            //print delivery Note by services
            setupFilterAndFavoritesPage.GoPrintDeliveryNoteByService();
            PrintReportPage reportPage1 = setupFilterAndFavoritesPage.PrintItemGenerique2(setupFilterAndFavoritesPage);
            reportPage1.Close();
            reportPage1.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            found = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi1 = new FileInfo(found);
            fi1.Refresh();
            Assert.IsTrue(fi1.Exists, $"{found} non généré");
            PdfDocument document1 = PdfDocument.Open(fi1.FullName);
            Page page11 = document1.GetPage(1);
            IEnumerable<Word> words1 = page11.GetWords();
            List<string> wordStrings1 = words1.Select(word => word.Text).ToList();
            string[] result1 = deliveryRName.Split(' ');
            bool IsRetourLigneEffectuer1 = setupFilterAndFavoritesPage.RetourLigneEffectuer(wordStrings1, result1);
            Assert.IsTrue(IsRetourLigneEffectuer1, "il faudrait espacer suffisamment pour que les lignes delivery note by services se chevauchent pas.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_Pagination()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            SetupPage setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.StartDate, DateUtils.Now.AddMonths(-10));
            setupFilterAndFavoritesPage.PageSize("8");
            int totalItems = setupFilterAndFavoritesPage.GetNameProvidersList().Count;
            bool isLessThan = totalItems <= 8;
            Assert.IsTrue(isLessThan, "Paggination ne fonctionne pas..");
            setupFilterAndFavoritesPage.PageSize("16");
            isLessThan = totalItems <= 16;
            Assert.IsTrue(isLessThan, "Paggination ne fonctionne pas..");
            setupFilterAndFavoritesPage.PageSize("30");
            isLessThan = totalItems <= 30;
            Assert.IsTrue(isLessThan, "Paggination ne fonctionne pas..");
            setupFilterAndFavoritesPage.PageSize("50");
            isLessThan = totalItems <= 50;
            Assert.IsTrue(isLessThan, "Paggination ne fonctionne pas..");
            setupFilterAndFavoritesPage.PageSize("100");
            isLessThan = totalItems <= 100;
            Assert.IsTrue(isLessThan, "Paggination ne fonctionne pas..");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FilterFAF_MakeFavRename()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string favouriteName = "testSetup" + new Random().Next(10, 99).ToString();
            string newFavouriteName = "testSetupUpdated" + new Random().Next(10, 99).ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();

            try
            {
                // Paramétrage et filtrage d'un favori
                setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);

                // Creartion de favorit 
                setupFilterAndFavoritesPage.MakeFavorite(favouriteName);
                setupFilterAndFavoritesPage.SetFavoriteText(favouriteName);
                //Assert
                Assert.IsTrue(setupFilterAndFavoritesPage.IsFavoritePresent(favouriteName), "Le favori " + favouriteName + " n'a pas été créé.");

                EditFavoriteModal editFavoriteModal = setupFilterAndFavoritesPage.EditFavoriteName(favouriteName);
                editFavoriteModal.RenameFavorite(newFavouriteName);
                editFavoriteModal.SaveEdit();
                setupFilterAndFavoritesPage.SetFavoriteText(newFavouriteName);
                string nouveauNomFavRecupere = setupFilterAndFavoritesPage.GetFavoriteNameText();
                // Assert
                Assert.IsTrue(newFavouriteName.Equals(nouveauNomFavRecupere), "Le nom du filtre favori n'a pas été renommé.");
            }
            finally
            {
                setupFilterAndFavoritesPage.DeleteFavorite(newFavouriteName);
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FAF_DONE()
        {
            // Préparer : Charger les informations nécessaires pour le test
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string expectedUrl = "/Production/Setup/SetupResults"; // URL cible attendue


            // Étape 1 : login
            HomePage homePage = LogInAsAdmin();

            // Étape 2 : Naviguer vers la page de configuration
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();

            // Étape 3 : Réinitialiser les filtres existants pour garantir un état propre
            setupFilterAndFavoritesPage.ResetFilters();

            // Étape 4 : Appliquer un filtre spécifique au site
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);

            // Étape 5 : Cliquer sur le bouton "Done" pour valider
            SetupPage setupPage = setupFilterAndFavoritesPage.Validate();

            // Étape 6 : Vérifier que la redirection vers les résultats est correcte
            string actualUrl = setupPage.GetCurrentUrl(); // Méthode pour obtenir l'URL actuelle

            //Assert
            Assert.IsTrue(actualUrl.Contains(expectedUrl), $"The URL does not contain the expected path: {expectedUrl}");

            // Étape 7 : Vérifier les données affichées sur la page des résultats
            Assert.IsTrue(setupPage.HasResults(), "Aucun résultat affiché sur la page des résultats.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_SearchByDeliveryRound()
        {
            //Prepare 
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            //Arrange 
            HomePage homePage = LogInAsAdmin();
            //ACT
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();

            setupFilterAndFavoritesPage.ResetFilters();
            //Filtrer pour chercher le deliveryround  créé et valider dans le dispatch 
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.SearchDeliveryRoundName, deliveryRound2Name);
            SetupPage setupPage = setupFilterAndFavoritesPage.Validate();
            string deliveryRoundName = setupPage.GetDeliveryRoundName();
            //assert 
            Assert.IsTrue(deliveryRoundName.Contains(deliveryRound2Name), "le deliveryound n'existe pas dans la page resultat");

        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FilterDeliveries()
        {
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            HomePage homePage = LogInAsAdmin();
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, yesterday);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
            SetupPage setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.Deliveries, delivery3Name);
            bool exist = setupPage.isFirstDeliveryRoundNameExist(delivery3Name);
            Assert.IsTrue(exist, "Le filtre ne fonctionne pas");
        }

        //[TestMethod]
        //[Timeout(Timeout)]
        //public void PR_SETUP_DeliveryRoundSingle()
        //{
        //    string DocFileNameZipBegin = "All_files_";
        //    string DocFileNamePdfBegin = "Delivery Note Recipes";
        //    string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
        //                string siteMAD = TestContext.Properties["Production_Site2"].ToString();
        //    string siteACE = TestContext.Properties["Production_Site1"].ToString();
        //    HomePage homePage = LogInAsAdmin();
        //    SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
        //    setupFilterAndFavoritesPage.ResetFilters();
        //    setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.StartDate, yesterday);
        //    setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteMAD);
        //    SetupPage setupPage = setupFilterAndFavoritesPage.Validate();
        //    setupFilterAndFavoritesPage.SelectDeliveryRound();
        //    setupFilterAndFavoritesPage.ClearDownloads();
        //    DeleteAllFileDownload();
        //    setupFilterAndFavoritesPage.GoPrintDeliveryRoundSingle();
        //    PrintReportPage reportPage = setupFilterAndFavoritesPage.PrintArchive(setupFilterAndFavoritesPage);
        //    reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
        //    reportPage.Close();
        //    string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
        //    FileInfo fichier = new FileInfo(trouve);
        //    fichier.Refresh();
        //    Assert.IsTrue(fichier.Exists, "Pas de print pdf");

        //}

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FilterWorkshop()
        {
            string workshop1 = "Ensamblaje";
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            var homePage = LogInAsAdmin();
            //act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.SearchDeliveryRoundName, deliveryRound2Name);
            setupPage.Filter(SetupPage.FilterType.Workshops, workshop1);
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            var deliveryRoundRecipeNameQties = setupDeliveryRoundTabPage.GetMenuNameAndQtyForOneDeliveryRound();
            Assert.IsTrue(deliveryRoundRecipeNameQties.DeliveryRoundName.Contains(deliveryRound2Name), String.Format(MessageErreur.FILTRE_ERRONE, "Workshop", workshop1));
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_SelectFavorite()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();

            //Filter setup
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.SetFavoriteText(favouriteName);
            setupFilterAndFavoritesPage.SelectFavorite();

            //Assert
            bool favouritePresent = setupFilterAndFavoritesPage.IsFavoritePresent(favouriteName);
            Assert.IsTrue(favouritePresent, "Le favori " + favouriteName + " n'est pas trouvé.");

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FilterDate()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();

            //Filter par date 
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.StartDate, yesterday);

            //Assert
            bool isPresentDelivery = setupPage.IsPresent();
            Assert.IsTrue(isPresentDelivery, "le filtre par date n'est pas fonctionnel.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FilterFAF_MakeFavorite()
        {
            //prepare
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            //arrange
            var homePage = LogInAsAdmin();
            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.WaitLoading();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.CustomerTypes, customerType1);
            //Create Favorite
            setupFilterAndFavoritesPage.MakeFavorite(favouriteName);
            setupFilterAndFavoritesPage.SetFavoriteText(favouriteName);
            //Assert
            bool favouritePresent = setupFilterAndFavoritesPage.IsFavoritePresent(favouriteName);
            Assert.IsTrue(favouritePresent, String.Format(MessageErreur.FILTRE_ERRONE, "FAVORITES", favouriteName));
            setupFilterAndFavoritesPage.SelectFavorite();
            string selectedCustomerType = setupFilterAndFavoritesPage.GetSelectedCustomerType();
            Assert.AreEqual(customerType1, selectedCustomerType, "Error: Le favoris est ajouté, il apparait dans les résultats de recherche des favoris");

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FAF_ResetFilter()
        {
            //prepare
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            //arrange
            var homePage = LogInAsAdmin();
            //Act
            var setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.CustomerTypes, customerType1);
            setupFilterAndFavoritesPage.ResetFilters();
            string siteText = setupFilterAndFavoritesPage.GetSitesAfterReset();
            bool siteApresResetFilter = siteText.Contains("Select a site...");
            Assert.IsTrue(siteApresResetFilter, String.Format(MessageErreur.FILTRE_ERRONE, "Erreur : Après avoir cliqué sur 'Réinitialiser les filtres', les filtres n'ont pas été réinitialisés correctement."));
            string selectedCustomerType = setupFilterAndFavoritesPage.GetCustomerType();
            Assert.AreNotEqual(customerType1, selectedCustomerType, "Erreur : Après avoir cliqué sur 'Réinitialiser les filtres', les filtres n'ont pas été réinitialisés correctement.");
        }
             
        [TestMethod]
        [Timeout(Timeout)]
           public void PR_SETUP_SearchByRecipeName()
            {
            //Prepare 
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            //Arrange 
            HomePage homePage = LogInAsAdmin();
            //ACT
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();

            setupFilterAndFavoritesPage.ResetFilters();
            //Filtrer pour chercher le deliveryround  créé et valider dans le dispatch 
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.RecipeName, condrenReceta1Name);

            SetupPage setupPage = setupFilterAndFavoritesPage.Validate();
            string deliveryRound= setupPage.GetDeliveryRoundName();
            //assert 
            Assert.IsTrue(deliveryRound.Contains(condrenDeliveryName), "le deliveryound n'existe pas dans la page resultat");
            }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FilterCustomer()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();

            //Filter par customer 
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.StartDate, yesterday);
            setupPage.Filter(SetupPage.FilterType.Customers, customer2Name);

            //Assert
            bool isPresentDelivery = setupPage.IsPresent();
            Assert.IsTrue(isPresentDelivery, "le filtre par customer n'est pas fonctionnel.");
        }
        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FAF_GoToResults()
        {
            //Prepare
            string workshop1 = TestContext.Properties["Item_Workshop"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string favouriteName = "testSetup" + new Random().Next(100, 1000).ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            try
            {
                // Paramétrage et filtrage d'un favori
                setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
                setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Deliveries, delivery2Name);
                setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Workshops, workshop1);
                // Creation de favorit 
                setupFilterAndFavoritesPage.MakeFavorite(favouriteName);
                setupFilterAndFavoritesPage.SetFavoriteText(favouriteName);
                //Assert
                Assert.IsTrue(setupFilterAndFavoritesPage.IsFavoritePresent(favouriteName), "Le favori " + favouriteName + " n'a pas été créé.");
                
                // Get les filters au niveau de la page Result 
                ResultModal resultFavoriteModal = setupFilterAndFavoritesPage.ResultFavoriteName(favouriteName);
                // Get site and compare with site added
                string siteSelected = resultFavoriteModal.GetSelectedSiteFilter();
                Assert.AreEqual(siteACE, siteSelected, "les sites sélectionnés ne sont pas compatibles.");
                // Get deliverie and compare with delivery added
                bool deliverieSelected = resultFavoriteModal.VerifyDeliveriesWrapper(delivery2Name);
                Assert.IsTrue(deliverieSelected, "Le deliverie name sélectionné n'a pas été correctement selectionnés après le filtrage.");
                resultFavoriteModal.ClickDeliveryRound();
                resultFavoriteModal.WaitPageLoading();
                // Get workshop and compare with workshop added
                bool workshopSelected = resultFavoriteModal.VerifyWorkshopsWrapper(workshop1);
                Assert.IsTrue(workshopSelected, "Le workshops name sélectionné n'a pas été correctement selectionnés après le filtrage.");
            }
            finally
            {
                // Go to production setup
                homePage.GoToProduction_SetupPage();
                setupFilterAndFavoritesPage.ResetFilters();
                // filter with site and delete favorit added
                setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
                setupFilterAndFavoritesPage.SetFavoriteText(favouriteName);
                setupFilterAndFavoritesPage.DeleteFavorite(favouriteName);
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FilterFAF_MakeFavDelete()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string favouriteName = "testSetup" + new Random().Next(1000, 9999).ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();

            // Paramétrage et filtrage d'un favori
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);

            // Creartion de favorit 
            setupFilterAndFavoritesPage.MakeFavorite(favouriteName);
            setupFilterAndFavoritesPage.SetFavoriteText(favouriteName);
            //Assert
            Assert.IsTrue(setupFilterAndFavoritesPage.IsFavoritePresent(favouriteName), "Le favori " + favouriteName + " n'a pas été créé.");
            // Delete Favorite Created
            setupFilterAndFavoritesPage.DeleteFavorite(favouriteName);
            // Filter with Favorite Created after delete 
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.SetFavoriteText(favouriteName);
            // Assert
            Assert.IsFalse(setupFilterAndFavoritesPage.IsFavoritePresent(favouriteName), "Le favori " + favouriteName + " n'a pas été créé.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FilterCustomerType()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();

            //Filter par customer 
            setupFilterAndFavoritesPage.ResetFilters();
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            setupPage.Filter(SetupPage.FilterType.StartDate, yesterday);
            setupPage.Filter(SetupPage.FilterType.CustomerTypes, CUSTOMERTYPE);

            //Assert
            bool isPresentDelivery = setupPage.IsPresent();
            Assert.IsTrue(isPresentDelivery, "le filtre par customer n'est pas fonctionnel.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_SETUP_FilterMealType()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string service2Name = "Service PR ACE 2" + new Random().Next(1000, 9999).ToString();
            string Delivery= "Delivery PR ACE 2" + new Random().Next(1000, 9999).ToString();
            string qty = "10";
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string deliveryRound2Name = "DR PR ACE 2" + new Random().Next(1000, 9999).ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string taxType = TestContext.Properties["Production_ItemTaxType"].ToString();
            string prodUnit = TestContext.Properties["Production_ItemProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string group2 = TestContext.Properties["Production_ItemGroup2"].ToString();
            string workshop2 = TestContext.Properties["Production_Workshop2"].ToString();
            string supplier2 = TestContext.Properties["Production_ItemSupplier2"].ToString();
            string item2Name = "Item PR ACE 2" + new Random().Next(1000, 9999).ToString();
            string recipe2Name = "Recipe PR ACE 2" + new Random().Next(1000, 9999).ToString();
            string recipeType2 = TestContext.Properties["Production_RecipeType2"].ToString();
            string recipeCookingMode2 = TestContext.Properties["Production_CookingMode2"].ToString();
            string recipeWorkshop2 = TestContext.Properties["Production_Workshop2"].ToString();
            string recipePortion2 = "18";
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string newTotalPortion = "100";
            string menu2Name = "Menu PR ACE 2" + new Random().Next(1000, 9999).ToString();
            string mealVariant1 = TestContext.Properties["Production_Meal1"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();
         
            // Create New Service and customer Price
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(service2Name, null, null, serviceCategorie, null, serviceType);
            ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            serviceGeneralInformationsPage.SetProduced(true);
            ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
            ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customer, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
            servicePage = pricePage.BackToList();

            // Create a new delivery, assign the already created service, and add the quantity.
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryPage.WaitPageLoading();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(Delivery, "A.F.A. LANZAROTE", siteACE, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();
            deliveryLoadingPage.AddService(service2Name);
            deliveryLoadingPage.AddQuantity(qty);
            deliveryPage = deliveryLoadingPage.BackToList();

            // Create a new delivery Round, assign the already created delivery.
            DeliveryRoundPage deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            DeliveryRoundCreateModalpage deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            DeliveryRoundGeneralInformationPage deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRound2Name, "ACE - ACE", DateUtils.Now, DateUtils.Now.AddDays(+31));
            DeliveryRoundDeliveryPage deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(Delivery);
            deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            deliveryRoundPage.ResetFilter();

            // Create a new item, and packaging.
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item2Name, group2, workshop2, taxType, prodUnit);
            // 1 packaging pour le site ACE
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.setYield("100");
            itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier2);
            itemPage = itemGeneralInformationPage.BackToList();

            // Create a new Recipe, assign the already created item, and variant.
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipe2Name, recipeType2, recipePortion2, true, null, recipeCookingMode2, recipeWorkshop2);
            // 1 variante pour ACE
            recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(item2Name);
            recipeVariantPage.SetTotalPortion(newTotalPortion);
            recipesPage = recipeVariantPage.BackToList();

            // Create a new Menu, assign the already created service, and recipe.
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            menusCreateModalPage.AddService(service2Name);
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menu2Name, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, recipeVariantForACE);
            menuDayViewPage.AddRecipe(recipe2Name);
            menusPage = menuDayViewPage.BackToList();

            // Filter the dispatch by Delivery and validate in the three tabs: 'Provisional Quantity', 'Quantity to Produce', and 'Quantity to Invoice'.
            DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);
            dispatchPage.Filter(DispatchPage.FilterType.Search, Delivery);
            PrevisionalQtyPage previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToInvoice();
            dispatchPage.ValidateAll();
            
            SetupFilterAndFavoritesPage setupFilterAndFavoritesPage = homePage.GoToProduction_SetupPage();
            setupFilterAndFavoritesPage.ResetFilters();
            // Paramétrage et filtrage d'un favori
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Sites, siteACE);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.MealTypes, mealVariant1);
            setupFilterAndFavoritesPage.Filter(SetupFilterAndFavoritesPage.FilterType.Deliveries, Delivery);
            var setupPage = setupFilterAndFavoritesPage.Validate();
            var setupDeliveryRoundTabPage = setupPage.GoToSetupDeliveryRoundTab();
            setupDeliveryRoundTabPage.FoldUnfoldAll();
            //Assert
            var recipeNameDisplayed = setupDeliveryRoundTabPage.GetRecipeNameRow();
            Assert.IsTrue(recipeNameDisplayed.Contains(mealVariant1), "Dans l'onglet Setup, onglet Delivery Round, le nom de la recette ne contient pas le 'Meal Types' qui a déjà été filtré !");
        }
    }
}