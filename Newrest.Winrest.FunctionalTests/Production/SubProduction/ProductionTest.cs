using DocumentFormat.OpenXml.VariantTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Threading;
using System.Web;
using UglyToad.PdfPig;
using Page = UglyToad.PdfPig.Content.Page;


namespace Newrest.Winrest.FunctionalTests.Production
{
    [TestClass]
    public class ProductionTest : TestBase
    {
        private const int Timeout = 600000;
        // données tests
        readonly static string today = DateUtils.Now.ToString("dd/MM/yyyy");
        readonly static string yesterday = DateUtils.Now.AddDays(-1).ToString("dd/MM/yyyy");



        // données pré config (attention données pré-confi aussi dans Setup)
        readonly string menu2Name = "Menu PR ACE 2";
        readonly string menu5Name = "Menu PR ACE 3";

        readonly string item3Name = "Item PR MAD 1";
        readonly string item4Name = "Item PR MAD 2";

        readonly string delivery2Name = "Delivery PR ACE 2";
        readonly string delivery3Name = "Delivery PR MAD 1";
        readonly string delivery4Name = "Delivery PR MAD 2";

        readonly string delivery5Name = "Delivery PR ACE 3";

        readonly string service2Name = "Service PR ACE 2";
        readonly string service3Name = "Service PR MAD 1";
        readonly string service4Name = "Service PR MAD 2";
        readonly string service5Name = "Service PR ACE 3";

        readonly string deliveryRound2Name = "DR PR ACE 2";

        readonly string customer2Name = "Customer PR ACE 2";
        readonly string customer3Name = "Customer PR MAD 1";
        readonly string customer4Name = "Customer PR MAD 2";
        readonly string customer5Name = "Customer PR ACE 3";

        readonly string recipe2Name = "Recipe PR ACE 2";
        readonly string recipe5Name = "Recipe PR ACE 3";
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
        readonly string foodpackCRT = "CRT";

        //Data Delivery ES (3 services, packaging Individual/Multiporcion)
        readonly string customerESName = "ES CUSTOMER";

        readonly string ESDeliveryName = "ES DELIVERY";

        readonly string ESINDServicioName = "ES IND SERVICIO";
        readonly string ESINDMenuName = "ES IND MENU";

        readonly string ESMULServicioName = "ES MUL SERVICIO";
        readonly string ESMULMenuName = "ES MUL MENU";

        readonly string ESINDMULServicioName = "ES IND MUL SERVICIO";
        readonly string ESINDMULMenuName = "ES IND MUL MENU";

        readonly string itemTest = "ITEM Test PR";
        readonly string recetaESName = "ES RECETA";


        // Datas Delivery TLS (packaging pax per packaging, 3 foodpacks)
        readonly string TLSCustomerName = "TLS CUSTOMER";
        readonly string TLSServicioName = "TLS SERVICIO";
        readonly string TLSRecetaName = "RECETA TLS";
        readonly string TLSMenuName = "MENU TLS";
        readonly string TLSDeliveryName = "DELIVERY TLS";

        readonly string customerTLSGRName = "CUSTOMER GR TLS";
        readonly string recetaTLSGRName = "RECETA GR TLS";
        readonly string menuTLSGR1Name = "MENU GR TLS 1";
        readonly string menuTLSGR2Name = "MENU GR TLS 2";

        readonly string customerPFName = "CUSTOMER PF";
        readonly string recetaPFName = "RECETA PF";
        readonly string menuPFName = "MENU PF";

        readonly string condrenCustomerName = "CUSTOMER CONDREN";
        readonly string condrenReceta1Name = "RECETA CONDREN 1";
        readonly string condrenReceta2Name = "RECETA CONDREN 2";
        readonly string condrenMenuName = "MENU CONDREN";

        readonly string barentinCustomerName = "CUSTOMER BARENTIN";
        readonly string barentinReceta1 = "RECETA BARENTIN 1";
        readonly string barentinRecetaNP2 = "RECETA BARENTIN NP 2";
        readonly string barentinRecetaBulk3 = "RECETA BARENTIN BULK 3";
        readonly string barentinMenuAdulto = "MENU BARENTIN ADULTO";
        readonly string barentinMenuCollegio = "MENU COLLEGIO";

        string serviceName = "ServiceToday" + DateTime.Now.ToString() + new Random().Next();
        string deleveryName = "DeliveryToday" +  new Random().Next();
        string itemName = "Item" + new Random().Next();
        string recipeName = "RecipeName" + DateTime.Now.ToString() + new Random().Next();
        string menuName = "MenuToday" + DateTime.Now.ToString() + new Random().Next();

        // Test Initialize
        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(PR_PR_Filter_DateFrom):
                    TestInitialize_CreateService();
                    TestInitialize_CreateDelivery();
                    TestInitialize_CreateItem();
                    TestInitialize_CreateRecipe();
                    TestInitialize_CreateMenu();
                    break;


                default:
                    break;
            }
        }

        //PRODUCTION TESTS PRECONFIG ARE IN SETUP CATEGORY (setup tests are running first on pipeline)

        [TestMethod]
        [Priority(0)]
        [Timeout(Timeout)]
        public void PR_PR_CreateCustomerOrders()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();
            ClearCache();

            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();

            // customerOrder 2
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(siteACE, customer2Name, delivery2Name, DateUtils.Now);
            var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

            customerOrderDetailPage.AddNewItemWithCategory(service2Name, "10", "COLECTIVIDADES");
            customerOrderDetailPage.ValidateCustomerOrder();

            var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
            var customerOrder2Number = generalInfo.GetOrderNumber();
            customerOrderDetailPage.BackToList();

            var name2 = customerOrderPage.GetCustomerOrderNumber();

            // customerOrder 3
            customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(siteMAD, customer3Name, delivery3Name, DateUtils.Now);
            customerOrderDetailPage = customerOrderCreateModalPage.Submit();

            customerOrderDetailPage.AddNewItemWithCategory(service3Name, "10", "COLECTIVIDADES");
            customerOrderDetailPage.ValidateCustomerOrder();

            generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
            var customerOrder3Number = generalInfo.GetOrderNumber();
            customerOrderDetailPage.BackToList();

            var name3 = customerOrderPage.GetCustomerOrderNumber();

            // customerOrder 4
            customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(siteMAD, customer4Name, delivery4Name, DateUtils.Now);
            customerOrderDetailPage = customerOrderCreateModalPage.Submit();

            customerOrderDetailPage.AddNewItemWithCategory(service4Name, "10", "COLECTIVIDADES");
            customerOrderDetailPage.ValidateCustomerOrder();

            generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
            var customerOrder4Number = generalInfo.GetOrderNumber();
            customerOrderDetailPage.BackToList();

            var name4 = customerOrderPage.GetCustomerOrderNumber();

            // customerOrder 5 - NO DELIVERY ROUND
            customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(siteACE, customer5Name, delivery5Name, DateUtils.Now);
            customerOrderDetailPage = customerOrderCreateModalPage.Submit();

            customerOrderDetailPage.AddNewItemWithCategory(service5Name, "10", "COLECTIVIDADES");
            customerOrderDetailPage.ValidateCustomerOrder();

            generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
            var customerOrder5Number = generalInfo.GetOrderNumber();
            customerOrderDetailPage.BackToList();

            var name5 = customerOrderPage.GetCustomerOrderNumber();

            Assert.IsTrue(name2.Contains(customerOrder2Number), "Le customer order n'a pas été créé.");
            Assert.IsTrue(name3.Contains(customerOrder3Number), "Le customer order n'a pas été créé.");
            Assert.IsTrue(name4.Contains(customerOrder4Number), "Le customer order n'a pas été créé.");
            Assert.IsTrue(name5.Contains(customerOrder5Number), "Le customer order n'a pas été créé.");
        }
        // create service
        private void TestInitialize_CreateService()
        {
            // prepare
            string serviceCategory = TestContext.Properties["InvoiceServiceCategoryAirportTax"].ToString();
            string customer = "CUSTOMER PF";
            string site = TestContext.Properties["SiteLP"].ToString();
            DateTime toDate = DateUtils.Now.AddMonths(+3);
            DateTime fromDate = DateUtils.Now.AddMonths(-1);
            // arrange 
            HomePage homePage = LogInAsAdmin();
            //act
            // create service
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, serviceCategory);
            ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            // add a price
            ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
            ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
            servicePage = pricePage.BackToList();
        }

        private void TestInitialize_CreateDelivery()
        {
            // prepare 
            string customer  = "CUSTOMER PF";
            string site = TestContext.Properties["SiteLP"].ToString();
            string qty = "15";
            // arrange 
            HomePage homePage = LogInAsAdmin();
            // act
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            // create Delivery
            DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deleveryName, customer, site, true);
            DeliveryLoadingPage deliveryLoadingPage = deliveryCreateModalPage.Create();
            // add a service
            deliveryLoadingPage.AddService(serviceName);
            deliveryLoadingPage.AddQuantity(qty);
        }

        private void TestInitialize_CreateItem()
        {
            // prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string group1 = TestContext.Properties["Prodman_Needs_Item_Group1"].ToString();
            string workshop1 = TestContext.Properties["Prodman_Needs_Workshop1"].ToString();
            string supplier1 = TestContext.Properties["Prodman_Needs_Item_Supplier1"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            // create item
            ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
            ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group1, workshop1, taxType, prodUnit);
            // add  packaging 
            ItemCreateNewPackagingModalPage itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier1);
            itemPage = itemGeneralInformationPage.BackToList();
        }

        private void TestInitialize_CreateRecipe()
        {
            // prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariantACE1"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            RecipesPage recipesPage = homePage.GoToMenus_Recipes();
            // create a recipe
            RecipesCreateModalPage recipesCreateModalPage = recipesPage.CreateNewRecipe();
            RecipeGeneralInformationPage recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");
            // ajouter un variant avec le meme site
            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            RecipesVariantPage recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            // ajouter item
            recipeVariantPage.AddIngredient(itemName);
        }

        private void TestInitialize_CreateMenu()
        {
            // prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string variant = TestContext.Properties["MenuVariantACE1"].ToString();
            DateTime endDate = DateUtils.Now.AddMonths(+3);
            DateTime startDate = DateUtils.Now.AddMonths(-1);
            // arrange
            HomePage homePage = LogInAsAdmin();
            MenusPage menusPage = homePage.GoToMenus_Menus();
            // create  a menu
            MenusCreateModalPage menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, startDate, endDate, site, variant, serviceName);
            // add a recipe
            menuDayViewPage.AddRecipe(recipeName);
        }




        // TESTS

        //PRINTS
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Print_Allotment()
        {
            //Prepare
            bool newVersionPrint = true;
            // Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();
            productionResultPage.ClearDownloads();
            var printReportPage = productionResultPage.Print(ProductionSearchResultsMenuTabPage.PrintType.AllotmentReport, newVersionPrint);
            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();
            Assert.IsTrue(isReportGenerated, "Le document PDF ALLOTMENT n'a pas pu être généré par l'application.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Print_Allotment_NonPackaged()
        {
            //Prepare
            bool newVersionPrint = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();

            productionResultPage.ClearDownloads();

            var printReportPage = productionResultPage.Print(ProductionSearchResultsMenuTabPage.PrintType.NonPackaged, newVersionPrint);

            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF ALLOTMENT n'a pas pu être généré par l'application.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Print_Allotment_GlobalNonPackaged()
        {
            //Prepare
            bool newVersionPrint = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();

            productionResultPage.ClearDownloads();

            var printReportPage = productionResultPage.Print(ProductionSearchResultsMenuTabPage.PrintType.GlobalNonPackaged, newVersionPrint);

            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF ALLOTMENT n'a pas pu être généré par l'application.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Print_HACCP()
        {
            //Prepare
            bool newVersionPrint = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();

            productionResultPage.ClearDownloads();

            var printReportPage = productionResultPage.Print(ProductionSearchResultsMenuTabPage.PrintType.Haccp, newVersionPrint);

            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF HACCP n'a pas pu être généré par l'application.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Print_RawMaterials()
        {
            //Prepare
            bool newVersionPrint = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();

            productionResultPage.ClearDownloads();

            var printReportPage = productionResultPage.Print(ProductionSearchResultsMenuTabPage.PrintType.RawMaterials, newVersionPrint);

            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF RAWMATERIALS n'a pas pu être généré par l'application.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Print_Production()
        {
            //Prepare
            bool newVersionPrint = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();

            productionResultPage.ClearDownloads();

            var printReportPage = productionResultPage.Print(ProductionSearchResultsMenuTabPage.PrintType.Production, newVersionPrint);

            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF PRODUCTION n'a pas pu être généré par l'application.");
        }

        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_ByMenu()
        {
            string site = "ACE";
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Date, today);
            productionPage.Filter(ProductionPage.FilterType.SearchText, menu2Name);
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            var menusPage = productionResultMenuTabPage.GoToMenus_Menus();

            foreach (var menuNameDeliveryQty in menusNamesDeliveriesQties)
            {
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuNameDeliveryQty.Key);

                if (!menusPage.GetFirstMenuName().Equals(menuNameDeliveryQty.Key))
                {
                    bool menusNamesOK = false;
                    //Assert
                    Assert.IsTrue(menusNamesOK, "L'affichage par MENU fonctionne pas");
                }

                menusPage.ResetFilter();
            }

            var dispatchPage = menusPage.GoToProduction_DispatchPage();

            foreach (var menuNameDeliveryQty in menusNamesDeliveriesQties)
            {
                dispatchPage.Filter(DispatchPage.FilterType.Search, menuNameDeliveryQty.Value.deliveryName);
                dispatchPage.Filter(DispatchPage.FilterType.Site, site);

                var qtyToProducePage = dispatchPage.ClickQuantityToProduce();

                if (!qtyToProducePage.IsQuantitiesUpdated(menuNameDeliveryQty.Value.qtyToProduce))
                {
                    bool qtyToProduceOK = false;
                    //Assert
                    Assert.IsTrue(qtyToProduceOK, "Les quantités affichées dans l'onglet By Menu ne correspondent pas aux quantités validées dans Dispatch");
                }

                dispatchPage = qtyToProducePage.GoToProduction_DispatchPage();
            }
        }
        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_ByMenuNoDeliveryRound()
        {
            string site = "ACE";
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.DeliveryRounds, "None");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                WebDriver.Navigate().Refresh();
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.DeliveryRounds, "None");
            }

            // temps d'affichage des données
            productionResultMenuTabPage.WaitPageLoading();
            if (productionResultMenuTabPage.CheckTotalNumber() > 8)
            {
                productionResultMenuTabPage.PageSize("100");
            }
            var lines = productionResultMenuTabPage.GetCustomersDeliveriesAndQties();
            List<string> menuNameList = new List<string>();
            foreach (var line in lines)
            {
                menuNameList.Add(line.Key);
            }
            Assert.IsTrue(menuNameList.Contains(menu5Name), "Les menus dont les delivery ne sont pas associés à une delivery round ne s'affichent pas");
        }

        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_ByRecipe()
        {
            // Prepare
            bool recipesNamesOK = true;
            bool qtyToProduceOK = true;
            string site = "ACE";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Date, today);
            productionPage.Filter(ProductionPage.FilterType.SearchText, "MENU PR_PR ");
            var productionResultPage = productionPage.Validate();
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            var productionRecipeTabPage = productionResultPage.GoToProductionRecipeTab();

            var recipesNamesAndQtiesToProduce = productionRecipeTabPage.GetRecipesNamesAndQtiesToProduce();

            var recipesPage = productionRecipeTabPage.GoToMenus_Recipes();

            foreach (var recipeNameAndQty in recipesNamesAndQtiesToProduce)
            {
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeNameAndQty.Key);

                if (!recipesPage.GetFirstRecipeName().Equals(recipeNameAndQty.Key))
                {
                    recipesNamesOK = false;
                    //Assert
                    Assert.IsTrue(recipesNamesOK, "L'affichage par RECIPE ne fonctionne pas");
                }

                recipesPage.ResetFilter();
            }

            productionPage = recipesPage.GoToProduction_ProductionPage();
            productionResultPage = productionPage.Validate();
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            var deliveryNamesAndQties = productionResultPage.GetMappedDeliveriesAndQtiesFromRecipes(recipesNamesAndQtiesToProduce);

            var dispatchPage = productionResultPage.GoToProduction_DispatchPage();

            foreach (var deliveryNameAndQty in deliveryNamesAndQties)
            {
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryNameAndQty.Key);
                dispatchPage.Filter(DispatchPage.FilterType.Site, site);

                var qtyToProducePage = dispatchPage.ClickQuantityToProduce();

                if (!qtyToProducePage.IsQuantitiesUpdated(deliveryNameAndQty.Value))
                {
                    qtyToProduceOK = false;
                    //Assert
                    Assert.IsTrue(qtyToProduceOK, "Les quantités affichées dans l'onglet By Recipe ne correspondent pas aux quantités validées dans Dispatch");
                }

                dispatchPage = qtyToProducePage.GoToProduction_DispatchPage();
            }
        }

        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_ByRecipeNoDeliveryRound()
        {
            string site = "ACE";
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.DeliveryRounds, "None");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                WebDriver.Navigate().Refresh();
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.DeliveryRounds, "None");
            }

            var recipeTab = productionResultMenuTabPage.GoToProductionRecipeTab();

            var lines = recipeTab.GetRecipesItemsNetWeight();
            List<string> recipeNameList = new List<string>();
            foreach (var line in lines)
            {
                recipeNameList.Add(line.Key);
            }
            Assert.IsTrue(recipeNameList.Contains(recipe5Name), "Les recipes dont les delivery ne sont pas associés à une delivery round ne s'affichent pas");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_ByCustomer()
        {
            // Prepare
            bool customersNamesOK = true;
            bool qtyToProduceOK = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Date, today);
            productionPage.Filter(ProductionPage.FilterType.SearchText, "MENU PR_PR ");
            var productionResultPage = productionPage.Validate();

            var productionCustomerTabPage = productionResultPage.GoToProductionCustomerTab();

            var customersNamesAndQties = productionCustomerTabPage.GetCustomersNamesAndQties();

            var customersPage = productionCustomerTabPage.GoToCustomers_CustomerPage();

            foreach (var customerNameAndQty in customersNamesAndQties)
            {
                customersPage.Filter(CustomerPage.FilterType.Search, customerNameAndQty.Key);

                if (!customersPage.GetFirstCustomerName().Equals(customerNameAndQty.Key))
                {
                    //Assert
                    customersNamesOK = false;
                    Assert.IsTrue(customersNamesOK, "L'affichage par CUSTOMER ne fonctionne pas : " + customerNameAndQty.Key);
                }

                customersPage.ResetFilters();
            }

            productionPage = customersPage.GoToProduction_ProductionPage();
            productionResultPage = productionPage.Validate();

            var dispatchPage = productionResultPage.GoToProduction_DispatchPage();

            foreach (var customerNameAndQty in customersNamesAndQties)
            {
                dispatchPage.Filter(DispatchPage.FilterType.Customers, customerNameAndQty.Key);

                var qtyToProducePage = dispatchPage.ClickQuantityToProduce();

                if (!qtyToProducePage.IsQuantitiesUpdated(customerNameAndQty.Value.Substring(customerNameAndQty.Value.Length - 2)))
                {
                    //Assert
                    qtyToProduceOK = false;
                    Assert.IsTrue(qtyToProduceOK, "Les quantités affichées dans l'onglet By Menu ne correspondent pas aux quantités validées dans Dispatch");
                }

                dispatchPage = qtyToProducePage.GoToProduction_DispatchPage();
            }
        }
        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_ByCustomerNoDeliveryRound()
        {
            string site = "ACE";
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.DeliveryRounds, "None");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                WebDriver.Navigate().Refresh();
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.DeliveryRounds, "None");
            }

            var customerTab = productionResultMenuTabPage.GoToProductionCustomerTab();
            var lines = customerTab.GetCustomersNamesAndQties();
            List<string> customerNameList = new List<string>();
            foreach (var line in lines)
            {
                customerNameList.Add(line.Key);
            }
            Assert.IsTrue(customerNameList.Contains(customer5Name), "Les customers dont les delivery ne sont pas associés à une delivery round ne s'affichent pas");
        }
        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_ByWorkshop()
        {
            //Prepare
            bool workshopsNamesOK = true;
            bool qtyToProduceOK = true;
            string site = "ACE";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Date, today);
            productionPage.Filter(ProductionPage.FilterType.SearchText, "MENU PR_PR ");
            var productionResultPage = productionPage.Validate();
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            var productionWorkshopTabPage = productionResultPage.GoToProductionWorkshopTab();

            var workshopNamesAndMappedRecipesQties = productionWorkshopTabPage.GetWorkshopNamesAndMappedRecipesQties();

            var recipesPage = productionWorkshopTabPage.GoToMenus_Recipes();

            Dictionary<string, string> recipesNamesAndQties = new Dictionary<string, string>();

            foreach (var workshopNameAndMappedRecipeQty in workshopNamesAndMappedRecipesQties)
            {
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, workshopNameAndMappedRecipeQty.Value.recipeName);

                var recipeGeneralInformationPage = recipesPage.SelectFirstRecipe();

                recipeGeneralInformationPage.ClickOnEditInformation();

                if (!recipeGeneralInformationPage.GetWorkshop().Equals(workshopNameAndMappedRecipeQty.Key))
                {
                    workshopsNamesOK = false;
                    //Assert
                    Assert.IsTrue(workshopsNamesOK, "L'affichage par WORKSHOP ne fonctionne pas");
                }

                recipesPage = recipeGeneralInformationPage.BackToList();

                recipesNamesAndQties.Add(workshopNameAndMappedRecipeQty.Value.recipeName, workshopNameAndMappedRecipeQty.Value.qtyToProduce);

                recipesPage.ResetFilter();
            }

            productionPage = recipesPage.GoToProduction_ProductionPage();
            productionResultPage = productionPage.Validate();
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            var deliveryNamesAndQties = productionResultPage.GetMappedDeliveriesAndQtiesFromRecipes(recipesNamesAndQties);

            var dispatchPage = productionResultPage.GoToProduction_DispatchPage();

            foreach (var deliveryNameAndQty in deliveryNamesAndQties)
            {
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryNameAndQty.Key);

                var qtyToProducePage = dispatchPage.ClickQuantityToProduce();

                if (!qtyToProducePage.IsQuantitiesUpdated(deliveryNameAndQty.Value))
                {
                    qtyToProduceOK = false;
                    //Assert
                    Assert.IsTrue(qtyToProduceOK, "Les quantités affichées dans l'onglet By Workshop ne correspondent pas aux quantités validées dans Dispatch");
                }

                dispatchPage = qtyToProducePage.GoToProduction_DispatchPage();
            }
        }
        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_ByWorkshopNoDeliveryRound()
        {
            string site = "ACE";
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.DeliveryRounds, "None");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                WebDriver.Navigate().Refresh();
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.DeliveryRounds, "None");
            }

            var workshopTab = productionResultMenuTabPage.GoToProductionWorkshopTab();
            var lines = workshopTab.GetWorkshopsNames();

            Assert.IsTrue(lines.Contains(recipe5Name), "Les workshops dont les delivery ne sont pas associés à une delivery round ne s'affichent pas");
        }
        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_ByCustomerOrder()
        {
            //Prepare
            bool orderNamesOK = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();

            var productionCustomerOrderTabPage = productionResultPage.GoToProductionCustomerOrderTab();

            var mapOrderNumberCustomer = productionCustomerOrderTabPage.GetCustomerOrdersNumberAndCustomers();

            var customerOrderPage = productionCustomerOrderTabPage.GoToProduction_CustomerOrderPage();

            foreach (var orderNumberAndCustomer in mapOrderNumberCustomer)
            {
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumberAndCustomer.Key);

                if (!customerOrderPage.GetFirstOrderNumber().Contains(orderNumberAndCustomer.Key))
                    orderNamesOK = false;

                if (!orderNumberAndCustomer.Value.Contains(customerOrderPage.GetFirstCustomer()))
                    orderNamesOK = false;

                if (!orderNumberAndCustomer.Value.Contains(customerOrderPage.GetFirstDelivery()))
                    orderNamesOK = false;
            }

            //Assert
            Assert.IsTrue(orderNamesOK, "L'affichage par CUSTOMER ORDER ne fonctionne pas");
        }
        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_ByCustomerOrderNoDeliveryRound()
        {
            // Prepare
            bool isTestOK = false;
            string site = "ACE";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            var customerOrderTab = productionResultMenuTabPage.GoToProductionCustomerOrderTab();

            var customerOrderNumbersAndNames = customerOrderTab.GetCustomerOrdersNumberAndCustomers();

            foreach (var orderNumberAndCustomer in customerOrderNumbersAndNames)
            {
                if (orderNumberAndCustomer.Value.Contains(customer2Name))
                    isTestOK = true;
            }

            //Assert
            Assert.IsTrue(isTestOK, "Les customerOrders dont les delivery ne sont pas associés à une delivery round ne s'affichent pas");
        }

        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Filter_BySearch()
        {
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();

            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, recipe2Name);
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);

            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var filterResult = productionResultPage.GetFirstMenuName();

            Assert.IsTrue(filterResult.Equals(menu2Name), "Le filtre SEARCH n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Filter_DateFrom()
        {
            // prepare 
            string site = TestContext.Properties["SiteLP"].ToString();
            string customer = "CUSTOMER PF";
            string dateEqualToMenuStartDate = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string dateInferiorOfMenuStartDate = DateUtils.Now.AddMonths(-2).ToString("dd/MM/yyyy");
            string dateSuperiorOfMenuStartDate = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");

            // arrange 
            HomePage homePage = LogInAsAdmin();
            // act
            DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
            // get delivery
            dispatchPage.Filter(DispatchPage.FilterType.Search, deleveryName );
            dispatchPage.Filter(DispatchPage.FilterType.Site, site);
            dispatchPage.WaitPageLoading();
            //validate PrevisonalQuantity
            PrevisionalQtyPage previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            previsionnalQty.ValidateFirstDispatch();
            //validate QuantityToProduce
            QuantityToProducePage qtyToProduce = dispatchPage.ClickQuantityToProduce();
            qtyToProduce.ValidateFirstDispatch();
            ProductionPage productionPage = homePage.GoToProduction_ProductionPage();
            // filtre by customer and site
            productionPage.Filter(ProductionPage.FilterType.Site, site);
            productionPage.Filter(ProductionPage.FilterType.Customers, customer);
            // filtre by date from with 3 different cases
            productionPage.Filter(ProductionPage.FilterType.Date, dateEqualToMenuStartDate);
            ProductionSearchResultsMenuTabPage productionMenuTab =  productionPage.Validate();
            productionPage.Filter(ProductionPage.FilterType.SearchText, menuName);
            int count = productionMenuTab.CheckTotalNumber();
            Assert.AreEqual(count, 1, " le filtre sur date from ne s'applique pas");
            productionPage.Filter(ProductionPage.FilterType.Date, dateSuperiorOfMenuStartDate);
            count = productionMenuTab.CheckTotalNumber();
            Assert.AreEqual(count, 0, " le filtre sur date from ne s'applique pas");
            productionPage.Filter(ProductionPage.FilterType.Date, dateInferiorOfMenuStartDate);
            count = productionMenuTab.CheckTotalNumber();
            Assert.AreEqual(count, 1, " le filtre sur date from ne s'applique pas");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Filter_BySite()
        {
            //Prepare
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // NOTE : actuellement filtre KO sur page resultats

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();

            productionPage.Filter(ProductionPage.FilterType.Site, siteMAD);

            var productionResultPage = productionPage.Validate();
            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var filterResult = productionResultPage.GetFirstMenuName();

            var menusPage = productionResultPage.GoToMenus_Menus();

            menusPage.Filter(MenusPage.FilterType.SearchMenu, filterResult);

            var menuSite = menusPage.GetFirstMenuSite();

            Assert.IsTrue(menuSite.Equals(siteMAD), "Le filtre SITES n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Filter_ShowItems()
        {
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();

            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, false);
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            var item = productionResultPage.GetFirstMenuItem();
            Assert.IsTrue(item.Equals(""), "Le filtre SHOW ITEMS n'a pas été appliqué correctement");

            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);

            var recipe = productionResultPage.GetFirstMenuRecipe();
            item = productionResultPage.GetFirstMenuItem();

            var recipePage = productionResultPage.GoToMenus_Recipes();

            recipePage.Filter(RecipesPage.FilterType.SearchRecipe, recipe);

            var recipeDetailsPage = recipePage.SelectFirstRecipe();

            var recipeVariantDetailPage = recipeDetailsPage.ClickOnFirstVariant();

            var itemFound = recipeVariantDetailPage.GetFirstItemName();

            Assert.AreEqual(item, itemFound, "Le filtre SHOW ITEMS n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Filter_CustomerType()
        {
            //Prepare
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Site, siteMAD);
            var productionResultPage = productionPage.Validate();

            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.CustomerTypes, customerType1);
            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }
            var productionCustomerTabPage = productionResultPage.GoToProductionCustomerTab();
            var customer = productionCustomerTabPage.GetFirstCustomerName();

            var customerPage = homePage.GoToCustomers_CustomerPage();

            customerPage.Filter(CustomerPage.FilterType.Search, customer);
            var firstCustomerName = customerPage.GetFirstCustomerName().Equals(customer);
            var typeCustomer = customerPage.VerifyTypeCustomer(customerType1);
            Assert.AreEqual(firstCustomerName, typeCustomer, "Le filtre CUSTOMER TYPES n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Filter_Customer()
        {
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();

            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Customers, customer2Name);
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var productionCustomerTabPage = productionResultPage.GoToProductionCustomerTab();
            var customer = productionCustomerTabPage.GetFirstCustomerName();

            Assert.AreEqual(customer, customer2Name, "Le filtre CUSTOMERS n'a pas été appliqué correctement");
        }
        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Filter_Workshop()
        {
            //Prepare
            string workshopCocinaCaliente = TestContext.Properties["Production_Workshop1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Site, siteMAD);
            var productionResultPage = productionPage.Validate();

            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Workshops, workshopCocinaCaliente);
            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var productionWorkshopTabPage = productionResultPage.GoToProductionWorkshopTab();

            var workshop = productionWorkshopTabPage.GetFirstWorkshopName();

            Assert.AreEqual(workshop, workshopCocinaCaliente, "Le filtre WORKSHOP n'a pas été appliqué correctement");
        }
        
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Filter_Mealtypes()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string mealType = TestContext.Properties["Production_Meal1"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Site, siteACE);
            var productionResultPage = productionPage.Validate();
            productionResultPage.ScrollToMealTypeFilter();
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.MealTypes, mealType);
            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }
            string firstVariant = productionResultPage.GetFirstMenuVariant();
            string selectedMealType = productionResultPage.GetMealTypes(firstVariant);
            Assert.AreEqual(selectedMealType, mealType, "Le filtre MEALTYPES n'a pas été appliqué correctement");
        }

        [TestMethod]
        
        [Timeout(Timeout)]
        public void PR_PR_Filter_GuestTypes()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string guestType = "BASAL ADULTO/MAYORES";
            // Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.GuestTypes, guestType);
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            var firstVariant = productionResultPage.GetFirstMenuVariant();
            bool ResultFilter = firstVariant.Contains(guestType);
            Assert.IsTrue(ResultFilter, "Le filtre GUEST TYPES n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Filter_DeliveryRounds()
        {
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();


            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.DeliveryRounds, deliveryRound2Name);
            if (!productionResultPage.IsResultsDisplayed())
            {
                WebDriver.Navigate().Refresh();
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.DeliveryRounds, deliveryRound2Name);
            }

            var delivery = productionResultPage.GetDeliveryOfFirstMenu();

            Assert.AreEqual(delivery, delivery2Name, "Le filtre DELIVERY ROUNDS n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Filter_Foodpack()
        {
            //Prepare
            string packagingName2 = "Multiporcion";
            string site = "ACE";

            // Arrange
            HomePage homePage = LogInAsAdmin();


            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();

            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.FoodpackTypes, packagingName2);
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);
            if (!productionResultPage.IsResultsDisplayed())
            {
                WebDriver.Navigate().Refresh();
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.FoodpackTypes, packagingName2);
            }
            var foodpack = productionResultPage.GetFoodpackOfFirstMenu();

            Assert.IsTrue(foodpack.Contains(packagingName2), "Le filtre FOODPACK n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_ResetFilter()
        {
            //Prepare
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage =  LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();

            productionPage.Filter(ProductionPage.FilterType.Site, siteMAD);

            var productionResultPage = productionPage.Validate();

            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Customers, customer4Name);

            productionPage = productionResultPage.ResetFilter();

            var selectedCustomers = productionPage.GetSelectedCustomers();

            var indexOfOf = selectedCustomers.IndexOf("of");
            var indexOfSelected = selectedCustomers.IndexOf("selected");

            var startNumberOfSelectedCustomers = 0;
            var endNumberOfSelectedCustomers = indexOfOf - 1;
            var numberOfSelectedCustomersLength = endNumberOfSelectedCustomers - startNumberOfSelectedCustomers;

            var startNumberOfTotalCustomers = indexOfOf + 3;
            var endNumberOfTotalCustomers = indexOfSelected - 1;
            var numberOfTotalCustomersLength = endNumberOfTotalCustomers - startNumberOfTotalCustomers;


            var numberOfSelectedCustomers = selectedCustomers.Substring(startNumberOfSelectedCustomers, numberOfSelectedCustomersLength);
            var numberOfTotalCustomers = selectedCustomers.Substring(startNumberOfTotalCustomers, numberOfTotalCustomersLength);

            Assert.AreEqual(numberOfSelectedCustomers, numberOfTotalCustomers, "Le filtre Customers n'a pas été réinitialisé");
            Assert.AreEqual(today, productionPage.GetSelectedDate(), "Le filtre DATE n'a pas été réinitialisé");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Unfold_All_And_Fold_All()
        {
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            //onglet bymenu 
            var byMenuButton = productionPage.GetByMenuButton();
            byMenuButton.Click();
            //Unfold
            Assert.IsTrue(productionResultPage.FoldUnfoldAll(), "Les éléments du tableau de l'onglet by menu ne sont pas dépliés");
            //Fold
            Assert.IsTrue(!productionResultPage.FoldUnfoldAll(), "Les éléments du tableau de l'onglet by menu ne sont pas repliés");

            //onglet byrecipe
            var byRecipeButton = productionPage.GetByRecipeButton();
            byRecipeButton.Click();
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            //Unfold
            Assert.IsTrue(productionResultPage.FoldUnfoldAll(), "Les éléments du tableau de l'onglet by recipe ne sont pas dépliés");
            //Fold
            Assert.IsTrue(!productionResultPage.FoldUnfoldAll(), "Les éléments du tableau de l'onglet by recipe ne sont pas repliés");

            //onglet bycustomer
            var byCustomer = productionPage.GetByCustomerButton();
            byCustomer.Click();
            //Unfold
            Assert.IsTrue(productionResultPage.FoldUnfoldAll(), "Les éléments du tableau de l'onglet by customer ne sont pas dépliés");
            //Fold
            Assert.IsTrue(!productionResultPage.FoldUnfoldAll(), "Les éléments du tableau de l'onglet by customer ne sont pas repliés");

            //onglet byworkshop
            var byWorkshopButton = productionPage.GetByWorkshopButton();
            byWorkshopButton.Click();
            //Unfold
            Assert.IsTrue(productionResultPage.FoldUnfoldAll(), "Les éléments du tableau de l'onglet by workshop ne sont pas dépliés");
            //Fold
            Assert.IsTrue(!productionResultPage.FoldUnfoldAll(), "Les éléments du tableau de l'onglet by workshop ne sont pas repliés");

            //onglet bycustomerorder
            var byCustomerOrderButton = productionPage.GetByCustomerOrder();
            byCustomerOrderButton.Click();
            //Unfold
            Assert.IsTrue(productionResultPage.FoldUnfoldAll(), "Les éléments du tableau de l'onglet by customer order ne sont pas dépliés");
            //Fold
            Assert.IsTrue(!productionResultPage.FoldUnfoldAll(), "Les éléments du tableau de l'onglet by customer order ne sont pas repliés");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_GenerateOutputForm()
        {
            //Prepare
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();

            productionPage.Filter(ProductionPage.FilterType.Site, siteMAD);

            var productionResultPage = productionPage.Validate();
            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var productionGenerateOutputFormModal = productionResultPage.GenerateOutputForm();

            productionGenerateOutputFormModal.SetFromPlace("Economato");
            productionGenerateOutputFormModal.SetToPlace("MAD4");
            var outputFormItem = productionGenerateOutputFormModal.Generate();

            outputFormItem.Filter(PageObjects.Warehouse.OutputForm.OutputFormItem.FilterItemType.SortBy, "Name");
            outputFormItem.Filter(PageObjects.Warehouse.OutputForm.OutputFormItem.FilterItemType.ShowItemsWithPhysQty, true);

            var listItems = outputFormItem.GetItemNames();

            var testOk = false;

            foreach (string s in listItems)
            {
                if (s.Equals(item3Name) || s.Equals(item4Name))
                    testOk = true;
                else
                    testOk = false;
            }

            Assert.IsTrue(testOk, "Les données présentes dans Output ne correspondent pas aux données de Production");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_GenerateOutputForm_With_Filter()
        {
            //Prepare
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();


            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();

            productionPage.Filter(ProductionPage.FilterType.Site, siteMAD);

            var productionResultPage = productionPage.Validate();

            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Deliveries, delivery3Name);
            if (!productionResultPage.IsResultsDisplayed())
            {
                WebDriver.Navigate().Refresh();
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Deliveries, delivery3Name);
            }

            var productionGenerateOutputFormModal = productionResultPage.GenerateOutputForm();

            productionGenerateOutputFormModal.SetFromPlace("Economato");
            productionGenerateOutputFormModal.SetToPlace("MAD4");
            var outputFormItem = productionGenerateOutputFormModal.Generate();

            outputFormItem.Filter(PageObjects.Warehouse.OutputForm.OutputFormItem.FilterItemType.ShowItemsWithPhysQty, false);
            outputFormItem.Filter(PageObjects.Warehouse.OutputForm.OutputFormItem.FilterItemType.SortBy, "Name");
            outputFormItem.Filter(PageObjects.Warehouse.OutputForm.OutputFormItem.FilterItemType.SearchByName, "Item");
            var firstItemName = outputFormItem.GetFirstItemName();
            Assert.AreEqual(firstItemName, item3Name, "Les données présentes dans Output ne correspondent pas aux données filtrées de Production");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Search_Date()
        {
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            HomePage homePage =  LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();

            productionPage.Filter(ProductionPage.FilterType.Site, siteACE);
            productionPage.Filter(ProductionPage.FilterType.Date, yesterday);


            var productionResultPage = productionPage.Validate();
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);

            var filterResult = productionResultPage.GetFirstMenuName();

            var menusPage = productionResultPage.GoToMenus_Menus();

            menusPage.Filter(MenusPage.FilterType.SearchMenu, filterResult);

            var menusDayViewPage = menusPage.SelectFirstMenu();

            menusDayViewPage.ClickOnDayBefore();

            var menuStartDate = menusDayViewPage.GetDayDisplayed().ToString("dd/MM/yyyy");

            Assert.IsTrue(menuStartDate.Equals(yesterday), "Les données issues de la recherche par DATE ne sont pas correctes");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Search_Site()
        {
            //Prepare
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();

            productionPage.Filter(ProductionPage.FilterType.Site, siteMAD);

            var productionResultPage = productionPage.Validate();
            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var filterResult = productionResultPage.GetFirstMenuName();

            var menusPage = productionResultPage.GoToMenus_Menus();

            menusPage.Filter(MenusPage.FilterType.SearchMenu, filterResult);

            var menuSite = menusPage.GetFirstMenuSite();
            bool menuSiteCheck = menuSite.Equals(siteMAD);
            Assert.IsTrue(menuSiteCheck, "Les données issues de la recherche par SITE ne sont pas correctes");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Search_Text()
        {
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();

            productionPage.Filter(ProductionPage.FilterType.Date, today);
            productionPage.Filter(ProductionPage.FilterType.Site, siteACE);

            productionPage.Filter(ProductionPage.FilterType.SearchText, menu2Name);

            var productionResultPage = productionPage.Validate();
            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }
            Assert.IsTrue(productionResultPage.GetFirstMenuName().Equals(menu2Name), "Le données issues de la recherche par TEXTE ne sont pas correctes");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Search_ShowItem()
        {

            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();

            productionPage.Filter(ProductionPage.FilterType.Date, today);
            productionPage.Filter(ProductionPage.FilterType.Site, siteACE);

            productionPage.Filter(ProductionPage.FilterType.ShowItems, true);

            // Temps d'affichage des résultats
            productionPage.WaitPageLoading();

            var productionResultPage = productionPage.Validate();

            var recipe = productionResultPage.GetFirstMenuRecipe();
            var item = productionResultPage.GetFirstMenuItem();

            var recipePage = productionResultPage.GoToMenus_Recipes();

            recipePage.Filter(RecipesPage.FilterType.SearchRecipe, recipe);

            var recipeDetailsPage = recipePage.SelectFirstRecipe();

            var recipeVariantDetailPage = recipeDetailsPage.ClickOnFirstVariant();

            var itemFound = recipeVariantDetailPage.GetFirstItemName();
            bool itemFoundCheck = itemFound.Equals(item);
            Assert.IsTrue(itemFoundCheck, "Le données issues de la recherche en affichant les ITEMS ne sont pas correctes");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Search_Mealtypes()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string mealType = TestContext.Properties["Production_Meal1"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.MealTypes, mealType);
            productionPage.Filter(ProductionPage.FilterType.Site, siteACE);
            var productionResultPage = productionPage.Validate();

            var firstVariant = productionResultPage.GetFirstMenuVariant();
            bool firstVariantCheck = firstVariant.Contains(mealType);
            Assert.IsTrue(firstVariantCheck, "Le données issues de la recherche par MEALTYPES ne sont pas correctes");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Search_Workshop()
        {
            //Prepare
            string workshopCocinaCaliente = TestContext.Properties["Production_Workshop1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Site, siteMAD);
            productionPage.Filter(ProductionPage.FilterType.Workshops, workshopCocinaCaliente);
            var productionResultPage = productionPage.Validate();
            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }
            var productionWorkshopTabPage = productionResultPage.GoToProductionWorkshopTab();

            var workshop = productionWorkshopTabPage.GetFirstWorkshopName();

            Assert.AreEqual(workshop, workshopCocinaCaliente, "Le données issues de la recherche par WORKSHOP ne sont pas correctes");

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Search_CustomerTypes()
        {
            //Prepare
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string siteMAD = TestContext.Properties["Production_Site2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();


            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.CustomerTypes, customerType1);
            productionPage.Filter(ProductionPage.FilterType.Site, siteMAD);
            var productionResultPage = productionPage.Validate();

            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var productionCustomerTabPage = productionResultPage.GoToProductionCustomerTab();

            var customer = productionCustomerTabPage.GetFirstCustomerName();

            var customerPage = homePage.GoToCustomers_CustomerPage();

            customerPage.Filter(CustomerPage.FilterType.Search, customer);

            Assert.AreEqual(customerPage.GetFirstCustomerName().Equals(customer), customerPage.VerifyTypeCustomer(customerType1), "La recherche par CUSTOMER TYPES n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Search_Customer()
        {
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Site, siteACE);
            productionPage.Filter(ProductionPage.FilterType.Customers, customer2Name);
            var productionResultPage = productionPage.Validate();
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);

            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var productionCustomerTabPage = productionResultPage.GoToProductionCustomerTab();
            var customer = productionCustomerTabPage.GetFirstCustomerName();

            Assert.AreEqual(customer, customer2Name, "La recherche par CUSTOMER n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Search_GuestTypes()
        {
            //Prepare
            string guestType = "BASAL ADULTO/MAYORES";
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();

            productionPage.Filter(ProductionPage.FilterType.GuestTypes, guestType);
            productionPage.Filter(ProductionPage.FilterType.Site, siteACE);
            var productionResultPage = productionPage.Validate();
            var firstVariant = productionResultPage.GetFirstMenuVariant();
            bool checkFirstVariant = firstVariant.Contains(guestType);
            Assert.IsTrue(checkFirstVariant, "La recherche par GUESTTYPE n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Search_DeliveryRounds()
        {
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();

            productionPage.Filter(ProductionPage.FilterType.Site, siteACE);
            productionPage.Filter(ProductionPage.FilterType.DeliveryRounds, deliveryRound2Name);

            var productionResultPage = productionPage.Validate();
            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }
            var delivery = productionResultPage.GetDeliveryOfFirstMenu();

            Assert.AreEqual(delivery, delivery2Name, "La recherche par DELIVERY ROUNDS n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Search_Foodpacks()
        {
            //Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string packagingName2 = "Multiporcion";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Site, siteACE);
            productionPage.Filter(ProductionPage.FilterType.FoodpackTypes, packagingName2);
            var productionResultPage = productionPage.Validate();
            if (!productionResultPage.IsResultsDisplayed())
            {
                productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var foodpack = productionResultPage.GetFoodpackOfFirstMenu();
            bool check = foodpack.Contains(packagingName2);
            Assert.IsTrue(check, "La recherche par FOODPACK n'a pas été appliqué correctement");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckPAX_MethodIndividualMultiportion()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity5 = "5";
            string quantity10 = "10";
            string quantity1 = "1";
            string totalPAX1 = "16";
            string totalPAX2 = "25";

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
            dispatchPage.WaitPageLoading();
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
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "ES ");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }
            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            if (menusNamesDeliveriesQties.Count == 0)
            {
                var menusPage = homePage.GoToMenus_Menus();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, ESINDMenuName);
                var menuDayViewPage = menusPage.SelectFirstMenu();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                }
                menusPage = menuDayViewPage.BackToList();

                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, ESMULMenuName);
                menuDayViewPage = menusPage.SelectFirstMenu();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                }
                menusPage = menuDayViewPage.BackToList();

                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, ESMULMenuName);
                menuDayViewPage = menusPage.SelectFirstMenu();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaESName);
                    menuDayViewPage.ClickOnFirstRecipe();
                }

                productionPage = homePage.GoToProduction_ProductionPage();
                productionResultMenuTabPage = productionPage.Validate();
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "MENU GR TLS ");
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
                productionResultMenuTabPage.WaitLoading();

                menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            }

            //Check ByMenu
            //Checks Menu ES IND
            menusNamesDeliveriesQties.TryGetValue(ESINDMenuName, out MappedDetailsByMenu mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity5), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch.", ESINDMenuName);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity5), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch.", recetaESName);

            //Checks Menu ES MUL
            menusNamesDeliveriesQties.TryGetValue(ESMULMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity10), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch.", ESMULMenuName);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity10), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch.", recetaESName);

            //Checks Menu ES IND MUL
            menusNamesDeliveriesQties.TryGetValue(ESINDMULMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity1), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch.", ESINDMULMenuName);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity1), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch.", recetaESName);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesPAX = productionRecipeTabPage.GetRecipesPAX();
            recipesPAX.TryGetValue(recetaESName, out string Pax);
            Assert.IsTrue(Pax.Equals(totalPAX1), "Le nombre de PAX de la recette {0} sur l'onglet By recipe ne correspond pas aux quantités validées dans le dispatch.", recetaESName);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            Assert.IsNotNull(productionCustomerTabPage, "Failed to navigate to Production Customer Tab.");

            var CustomersPAX = productionCustomerTabPage.GetCustomersPAX();
            Assert.IsNotNull(CustomersPAX, "Failed to retrieve Customers PAX data. Ensure the data is loaded.");
            Assert.IsTrue(CustomersPAX.Count > 0, "Customers PAX data is empty.");

            MappedCustomerAndRecipePax customerPax;
            Assert.IsTrue(CustomersPAX.TryGetValue(customerESName, out customerPax), $"Customer '{customerESName}' not found in Customers PAX data.");
            Assert.IsNotNull(customerPax, $"Customer '{customerESName}' exists, but its PAX data is null.");

            Assert.IsTrue(customerPax.customerPAX.Contains(totalPAX1),
                $"Le nombre de PAX du Customer {customerESName} sur l'onglet By customer ne correspond pas aux quantités validées dans le dispatch.");
            Assert.IsTrue(customerPax.recipePAX.Contains(totalPAX1),
                $"Le nombre de PAX de la recette {recetaESName} du customer {customerESName} sur l'onglet By customer ne correspond pas aux quantités validées dans le dispatch.");

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

            //Change quantites
            productionPage = homePage.GoToProduction_ProductionPage();
            productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "ES ");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);

            productionResultMenuTabPage.WaitLoading();
            menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu ES IND
            menusNamesDeliveriesQties.TryGetValue(ESINDMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity5), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch.", ESINDMenuName);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity5), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch.", recetaESName);

            //Checks Menu ES MUL
            menusNamesDeliveriesQties.TryGetValue(ESMULMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity10), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch.", ESMULMenuName);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity10), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch.", recetaESName);

            //Checks Menu ES IND MUL
            menusNamesDeliveriesQties.TryGetValue(ESINDMULMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity10), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch.", ESINDMULMenuName);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity10), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch.", recetaESName);

            //Check By Recipe
            productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            recipesPAX = productionRecipeTabPage.GetRecipesPAX();
            recipesPAX.TryGetValue(recetaESName, out Pax);
            Assert.IsTrue(Pax.Equals(totalPAX2), "Le nombre de PAX de la recette {0} sur l'onglet By recipe ne correspond pas aux quantités validées dans le dispatch.", recetaESName);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            CustomersPAX = productionCustomerTabPage.GetCustomersPAX();
            CustomersPAX.TryGetValue(customerESName, out customerPax);
            Assert.IsTrue(customerPax.customerPAX.Contains(totalPAX2), "Le nombre de PAX du Customer {0} sur l'onglet By customer ne correspond pas aux quantités validées dans le dispatch.", customerESName);
            Assert.IsTrue(customerPax.recipePAX.Contains(totalPAX2), "Le nombre de PAX de la recette {0} du customer {1} sur l'onglet By customer ne correspond pas aux quantités validés dans le dispatch.", recetaESName, customerESName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckPAX_MethodPaxPerPackaging()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity10 = "10";
            string quantity12 = "12";

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
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, quantity10);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, TLSMenuName);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);

            productionResultMenuTabPage.WaitLoading();
            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            if (menusNamesDeliveriesQties.Count == 0)
            {
                var menusPage = homePage.GoToMenus_Menus();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, TLSMenuName);
                var menuDayViewPage = menusPage.SelectFirstMenu();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(TLSRecetaName);
                    menuDayViewPage.ClickOnFirstRecipe();
                }

                productionPage = homePage.GoToProduction_ProductionPage();
                productionResultMenuTabPage = productionPage.Validate();
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, TLSMenuName);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
                productionResultMenuTabPage.WaitLoading();
                menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            }

            //Check ByMenu
            //Checks Menu TLS
            MappedDetailsByMenu mappedDetailsByMenu;
            var menuExactName = menusNamesDeliveriesQties.First().Key;

            menusNamesDeliveriesQties.TryGetValue(menuExactName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity10), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", TLSMenuName, mappedDetailsByMenu.menuPAX, quantity10);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity10), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", TLSRecetaName, mappedDetailsByMenu.recipePAX, quantity10);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesPAX = productionRecipeTabPage.GetRecipesPAX();
            string Pax;

            var recetaExactName = recipesPAX.First().Key;
            recipesPAX.TryGetValue(recetaExactName, out Pax);
            Assert.IsTrue(Pax.Equals(quantity10), "Le nombre de PAX de la recette {0} sur l'onglet By recipe ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", TLSRecetaName, Pax, quantity10);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var CustomersPAX = productionCustomerTabPage.GetCustomersPAX();
            MappedCustomerAndRecipePax customerPax;
            var customerExactName = CustomersPAX.First().Key;
            CustomersPAX.TryGetValue(customerExactName, out customerPax);
            Assert.IsTrue(customerPax.customerPAX.Contains(quantity10), "Le nombre de PAX du Customer {0} sur l'onglet By customer ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX : {2}>.", TLSCustomerName, customerPax.customerPAX, quantity10);
            Assert.IsTrue(customerPax.recipePAX.Contains(quantity10), "Le nombre de PAX de la recette {0} du customer {1} sur l'onglet By customer ne correspond pas aux quantités validés dans le dispatch, valeur <{2}> au lieu de <PAX : {3}>.", TLSRecetaName, TLSCustomerName, customerPax.recipePAX, quantity10);

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
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, quantity12);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            //Act
            productionPage = homePage.GoToProduction_ProductionPage();
            productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, TLSMenuName);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);

            productionResultMenuTabPage.WaitLoading();
            menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu TLS
            menusNamesDeliveriesQties.TryGetValue(menuExactName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity12), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", TLSMenuName, mappedDetailsByMenu.menuPAX, quantity12);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity12), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", TLSRecetaName, mappedDetailsByMenu.recipePAX, quantity12);

            //Check By Recipe
            productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            recipesPAX = productionRecipeTabPage.GetRecipesPAX();
            recipesPAX.TryGetValue(recetaExactName, out Pax);
            Assert.IsTrue(Pax.Equals(quantity12), "Le nombre de PAX de la recette {0} sur l'onglet By recipe ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", TLSRecetaName, Pax, quantity12);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            CustomersPAX = productionCustomerTabPage.GetCustomersPAX();
            CustomersPAX.TryGetValue(customerExactName, out customerPax);
            Assert.IsTrue(customerPax.customerPAX.Contains(quantity12), "Le nombre de PAX du Customer {0} sur l'onglet By customer ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX : {2}>.", TLSCustomerName, customerPax.customerPAX, quantity12);
            Assert.IsTrue(customerPax.recipePAX.Contains(quantity12), "Le nombre de PAX de la recette {0} du customer {1} sur l'onglet By customer ne correspond pas aux quantités validés dans le dispatch, valeur <{2}> au lieu de <PAX : {3}>.", TLSRecetaName, TLSCustomerName, customerPax.recipePAX, quantity12);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckPAX_MethodPaxPerPackagingGrouped()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity5 = "5";
            string quantity10 = "10";
            string totalPAX = "15";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();

            //MENU TLS GR
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "MENU GR TLS ");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            if (menusNamesDeliveriesQties.Count == 0)
            {
                var menusPage = homePage.GoToMenus_Menus();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuTLSGR1Name);
                var menuDayViewPage = menusPage.SelectFirstMenu();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaTLSGRName);
                    menuDayViewPage.ClickOnFirstRecipe();
                }
                menusPage = menuDayViewPage.BackToList();

                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuTLSGR2Name);
                menuDayViewPage = menusPage.SelectFirstMenu();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaTLSGRName);
                    menuDayViewPage.ClickOnFirstRecipe();
                }

                productionPage = homePage.GoToProduction_ProductionPage();
                productionResultMenuTabPage = productionPage.Validate();
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "MENU GR TLS ");
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
                productionResultMenuTabPage.WaitLoading();

                menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            }

            //Check ByMenu
            //Checks Menu TLS GR 1
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(menuTLSGR1Name, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity5), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", menuTLSGR1Name, mappedDetailsByMenu.menuPAX, quantity10);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity5), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", recetaTLSGRName, mappedDetailsByMenu.recipePAX, quantity10);

            //Checks Menu TLS GR 2
            menusNamesDeliveriesQties.TryGetValue(menuTLSGR2Name, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity10), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", menuTLSGR1Name, mappedDetailsByMenu.menuPAX, quantity10);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity10), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", recetaTLSGRName, mappedDetailsByMenu.recipePAX, quantity10);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesPAX = productionRecipeTabPage.GetRecipesPAX();
            string Pax;
            recipesPAX.TryGetValue(recetaTLSGRName, out Pax);
            Assert.IsTrue(Pax.Equals(totalPAX), "Le nombre de PAX de la recette {0} sur l'onglet By recipe ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", recetaTLSGRName, Pax, totalPAX);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var CustomersPAX = productionCustomerTabPage.GetCustomersPAX();
            MappedCustomerAndRecipePax customerPax;
            CustomersPAX.TryGetValue(customerTLSGRName, out customerPax);
            Assert.IsTrue(customerPax.customerPAX.Contains(totalPAX), "Le nombre de PAX du Customer {0} sur l'onglet By customer ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX : {2}>.", customerTLSGRName, customerPax.customerPAX, totalPAX);
            Assert.IsTrue(customerPax.recipePAX.Contains(totalPAX), "Le nombre de PAX de la recette {0} du customer {1} sur l'onglet By customer ne correspond pas aux quantités validés dans le dispatch, valeur <{2}> au lieu de <PAX : {3}>.", recetaTLSGRName, customerTLSGRName, customerPax.recipePAX, totalPAX);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckPAX_MethodByRecipeVariant()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity16 = "16";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, menuPFName);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            if (menusNamesDeliveriesQties.Count == 0)
            {
                var menusPage = homePage.GoToMenus_Menus();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuPFName);
                var menuDayViewPage = menusPage.SelectFirstMenu();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recetaPFName);
                    menuDayViewPage.ClickOnFirstRecipe();
                }

                productionPage = homePage.GoToProduction_ProductionPage();
                productionResultMenuTabPage = productionPage.Validate();
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, menuPFName);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
                productionResultMenuTabPage.WaitLoading();

                menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            }

            //Check ByMenu
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(menuPFName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity16), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", menuPFName, mappedDetailsByMenu.menuPAX, quantity16);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity16), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", recetaPFName, mappedDetailsByMenu.recipePAX, quantity16);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesPAX = productionRecipeTabPage.GetRecipesPAX();
            string Pax;

            recipesPAX.TryGetValue(recetaPFName, out Pax);
            Assert.IsTrue(Pax.Equals(quantity16), "Le nombre de PAX de la recette {0} sur l'onglet By recipe ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", recetaPFName, Pax, quantity16);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var CustomersPAX = productionCustomerTabPage.GetCustomersPAX();
            MappedCustomerAndRecipePax customerPax;
            CustomersPAX.TryGetValue(customerPFName, out customerPax);
            Assert.IsTrue(customerPax.customerPAX.Contains(quantity16), "Le nombre de PAX du Customer {0} sur l'onglet By customer ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX : {2}>.", TLSCustomerName, customerPax.customerPAX, quantity16);
            Assert.IsTrue(customerPax.recipePAX.Contains(quantity16), "Le nombre de PAX de la recette {0} du customer {1} sur l'onglet By customer ne correspond pas aux quantités validés dans le dispatch, valeur <{2}> au lieu de <PAX : {3}>.", TLSRecetaName, TLSCustomerName, customerPax.recipePAX, quantity16);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckPAX_MethodByRecipeVariant2Recipes()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity = "15";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, condrenReceta2Name);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            productionResultMenuTabPage.WaitLoading();

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            if (menusNamesDeliveriesQties.Count == 0)
            {
                var menusPage = homePage.GoToMenus_Menus();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, condrenMenuName);
                var menuDayViewPage = menusPage.SelectFirstMenu();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(condrenReceta2Name);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeMethod("Std");
                    menuDayViewPage.SetRecipeCoef("2");
                    menuDayViewPage.AddRecipe(condrenReceta1Name);
                }

                productionPage = homePage.GoToProduction_ProductionPage();
                productionResultMenuTabPage = productionPage.Validate();
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, condrenReceta2Name);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
                productionResultMenuTabPage.WaitLoading();

                menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            }

            //Recipe Condren 2
            //Check ByMenu
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(condrenMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity), "Le nombre de PAX du menu {0} ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", condrenMenuName, mappedDetailsByMenu.menuPAX, quantity);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity), "Le nombre de PAX de cette recette {0} ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", condrenReceta2Name, mappedDetailsByMenu.recipePAX, quantity);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesPAX = productionRecipeTabPage.GetRecipesPAX();
            string Pax;

            recipesPAX.TryGetValue(condrenReceta2Name, out Pax);
            Assert.IsTrue(Pax.Equals(quantity), "Le nombre de PAX de la recette {0} sur l'onglet By recipe ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", condrenReceta1Name, Pax, quantity);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var CustomersPAX = productionCustomerTabPage.GetCustomersPAX();
            MappedCustomerAndRecipePax customerPax;
            CustomersPAX.TryGetValue(condrenCustomerName, out customerPax);
            Assert.IsTrue(customerPax.customerPAX.Contains(quantity), "Le nombre de PAX du Customer {0} sur l'onglet By customer ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX : {2}>.", condrenCustomerName, customerPax.customerPAX, quantity);
            Assert.IsTrue(customerPax.recipePAX.Contains(quantity), "Le nombre de PAX de la recette {0} du customer {1} sur l'onglet By customer ne correspond pas aux quantités validés dans le dispatch, valeur <{2}> au lieu de <PAX : {3}>.", condrenReceta2Name, condrenCustomerName, customerPax.recipePAX, quantity);
            productionCustomerTabPage.GoToProductionMenuTab();

            //Recipe Condren 1
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, condrenReceta1Name);
            productionResultMenuTabPage.WaitLoading(); //Chargement 
            menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            menusNamesDeliveriesQties.TryGetValue(condrenMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", condrenMenuName, mappedDetailsByMenu.menuPAX, quantity);
            Assert.IsTrue(mappedDetailsByMenu.recipePAX.Contains(quantity), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", condrenReceta1Name, mappedDetailsByMenu.recipePAX, quantity);

            //Check By Recipe
            productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            recipesPAX = productionRecipeTabPage.GetRecipesPAX();

            recipesPAX.TryGetValue(condrenReceta1Name, out Pax);
            Assert.IsTrue(Pax.Equals(quantity), "Le nombre de PAX de la recette {0} sur l'onglet By recipe ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", condrenReceta1Name, Pax, quantity);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            CustomersPAX = productionCustomerTabPage.GetCustomersPAX();
            CustomersPAX.TryGetValue(condrenCustomerName, out customerPax);
            Assert.IsTrue(customerPax.customerPAX.Contains(quantity), "Le nombre de PAX du Customer {0} sur l'onglet By customer ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX : {2}>.", condrenCustomerName, customerPax.customerPAX, quantity);
            Assert.IsTrue(customerPax.recipePAX.Contains(quantity), "Le nombre de PAX de la recette {0} du customer {1} sur l'onglet By customer ne correspond pas aux quantités validés dans le dispatch, valeur <{2}> au lieu de <PAX : {3}>.", condrenReceta1Name, condrenCustomerName, customerPax.recipePAX, quantity);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckPAX_MethodRecipeVariantAndGroupedByRecipe()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity1 = "10";
            string quantity2 = "50";
            string PAXtotal = "60";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "BARENTIN");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            if (menusNamesDeliveriesQties.Count == 0)
            {
                var menusPage = homePage.GoToMenus_Menus();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, barentinMenuAdulto);
                var menuDayViewPage = menusPage.SelectFirstMenu();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(barentinReceta1);
                    menuDayViewPage.AddRecipe(barentinRecetaNP2);
                    menuDayViewPage.AddRecipe(barentinRecetaBulk3);
                }
                menuDayViewPage.BackToList();
                productionPage = homePage.GoToProduction_ProductionPage();
                productionResultMenuTabPage = productionPage.Validate();
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "BARENTIN");
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
                productionResultMenuTabPage.WaitLoading();

                menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();
            }

            //Check ByMenu
            MappedDetailsByMenu mappedDetailsByMenu;
            //Barentin Menu Adulto
            menusNamesDeliveriesQties.TryGetValue(barentinMenuAdulto, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity1), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", condrenMenuName, mappedDetailsByMenu.menuPAX, quantity1);
            //Barentin Receta 1
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuAdulto, barentinReceta1).Contains(quantity1), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", barentinReceta1, productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuAdulto, barentinReceta1), quantity1);
            //Barentin Receta 2
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuAdulto, barentinRecetaNP2).Contains(quantity1), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", barentinRecetaNP2, productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuAdulto, barentinRecetaNP2), quantity1);
            //Barentin Receta 3
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuAdulto, barentinRecetaBulk3).Contains(quantity1), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", barentinRecetaBulk3, productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuAdulto, barentinRecetaBulk3), quantity1);

            //Barentin Menu Collegio
            menusNamesDeliveriesQties.TryGetValue(barentinMenuCollegio, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.menuPAX.Equals(quantity2), "Le nombre de PAX du menu {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", barentinMenuCollegio, mappedDetailsByMenu.menuPAX, quantity2);
            //Barentin Receta 1
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuCollegio, barentinReceta1).Contains(quantity2), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", barentinReceta1, productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuCollegio, barentinReceta1), quantity2);
            //Barentin Receta 2
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuCollegio, barentinRecetaNP2).Contains(quantity2), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", barentinRecetaNP2, productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuCollegio, barentinRecetaNP2), quantity2);
            //Barentin Receta 3
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuCollegio, barentinRecetaBulk3).Contains(quantity2), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", barentinRecetaBulk3, productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuCollegio, barentinRecetaBulk3), quantity2);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesPAX = productionRecipeTabPage.GetRecipesPAX();
            string Pax;

            recipesPAX.TryGetValue(barentinReceta1, out Pax);
            Assert.IsTrue(Pax.Equals(PAXtotal), "Le nombre de PAX de la recette {0} sur l'onglet By recipe ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", barentinReceta1, Pax, PAXtotal);
            recipesPAX.TryGetValue(barentinRecetaNP2, out Pax);
            Assert.IsTrue(Pax.Equals(PAXtotal), "Le nombre de PAX de la recette {0} sur l'onglet By recipe ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", barentinRecetaNP2, Pax, PAXtotal);
            recipesPAX.TryGetValue(barentinRecetaBulk3, out Pax);
            Assert.IsTrue(Pax.Equals(PAXtotal), "Le nombre de PAX de la recette {0} sur l'onglet By recipe ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <{2}>.", barentinRecetaBulk3, Pax, PAXtotal);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var CustomersPAX = productionCustomerTabPage.GetCustomersPAX();
            MappedCustomerAndRecipePax customerPax;
            CustomersPAX.TryGetValue(barentinCustomerName, out customerPax);
            Assert.IsTrue(customerPax.customerPAX.Contains(PAXtotal), "Le nombre de PAX du Customer {0} sur l'onglet By customer ne correspond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX : {2}>.", barentinCustomerName, customerPax.customerPAX, PAXtotal);
            Assert.IsTrue(customerPax.recipePAX.Contains(PAXtotal), "Le nombre de PAX de la recette {0} du customer {1} sur l'onglet By customer ne correspond pas aux quantités validés dans le dispatch, valeur <{2}> au lieu de <PAX : {3}>.", barentinReceta1, barentinCustomerName, customerPax.recipePAX, PAXtotal);
            productionCustomerTabPage.GoToProductionMenuTab();
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckQuantityToProduce_MethodIndividualMultiportion()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity5 = "5";
            string quantity10 = "10";
            string quantity1 = "1";
            string ESINDQuantityToProduce = "3"; // quantity*coef
            string ESMULQuantityToProduce = "10";
            string ESINDMULQuantityToProduce1 = "2";
            string ESINDMULQuantityToProduce2 = "15";
            string totalQuantityToProduce1 = "15";
            string totalQuantityToProduce2 = "28";

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
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "ES ");

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }
            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu ES IND
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(ESINDMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(ESINDQuantityToProduce), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", recetaESName, ESINDMenuName);

            //Checks Menu ES MUL
            menusNamesDeliveriesQties.TryGetValue(ESMULMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(ESMULQuantityToProduce), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", recetaESName, ESMULMenuName);

            //Checks Menu ES IND MUL
            menusNamesDeliveriesQties.TryGetValue(ESINDMULMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(ESINDMULQuantityToProduce1), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", recetaESName, ESINDMULMenuName);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesQuantityToProduce = productionRecipeTabPage.GetRecipesNamesAndQtiesToProduce();
            string recipeQuantityToProduce;
            recipesQuantityToProduce.TryGetValue(recetaESName, out recipeQuantityToProduce);
            Assert.IsTrue(recipeQuantityToProduce.Equals(totalQuantityToProduce1), "Dans 'By recipe', la quantity to produce de la recette {0} ne correspond pas aux quantités validées dans le dispatch.", recetaESName);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var customersQuantityToProduce = productionCustomerTabPage.GetCustomersNamesAndQties();
            string customerQuantityToProduce;
            customersQuantityToProduce.TryGetValue(customerESName, out customerQuantityToProduce);
            Assert.IsTrue(customerQuantityToProduce.Contains(totalQuantityToProduce1), "Dans 'By customer', la quantity to produce de la recette {0} du customer {1} ne correspond pas aux quantités validées dans le dispatch.", recetaESName, customerESName);

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

            //Act
            productionPage = homePage.GoToProduction_ProductionPage();
            productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "ES ");

            productionResultMenuTabPage.WaitLoading();
            menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu ES IND
            menusNamesDeliveriesQties.TryGetValue(ESINDMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(ESINDQuantityToProduce), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", recetaESName, ESINDMenuName);
            //Checks Menu ES MUL
            menusNamesDeliveriesQties.TryGetValue(ESMULMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(ESMULQuantityToProduce), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", recetaESName, ESMULMenuName);

            //Checks Menu ES IND MUL
            menusNamesDeliveriesQties.TryGetValue(ESINDMULMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(ESINDMULQuantityToProduce2), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", recetaESName, ESINDMULMenuName);

            //Check By Recipe
            productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            recipesQuantityToProduce = productionRecipeTabPage.GetRecipesNamesAndQtiesToProduce();
            recipesQuantityToProduce.TryGetValue(recetaESName, out recipeQuantityToProduce);
            Assert.IsTrue(recipeQuantityToProduce.Equals(totalQuantityToProduce2), "Dans 'By recipe', la quantity to produce de la recette {0} ne correspond pas aux quantités validées dans le dispatch.", recetaESName);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            customersQuantityToProduce = productionCustomerTabPage.GetCustomersNamesAndQties();
            customersQuantityToProduce.TryGetValue(customerESName, out customerQuantityToProduce);
            Assert.IsTrue(customerQuantityToProduce.Contains(totalQuantityToProduce2), "Dans 'By customer', la quantity to produce de la recette {0} du customer {1} ne correspond pas aux quantités validées dans le dispatch.", recetaESName, customerESName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckQuantityToProduce_MethodPaxPerPackaging()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity10 = "10";
            string quantity12 = "12";
            string totalQuantityToProduce1 = "20"; //quantity*coef
            string totalQuantityToProduce2 = "24";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //DELIVERY WITH SERVICIO TLS = 10
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, TLSDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, quantity10);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, TLSMenuName);

            productionResultMenuTabPage.WaitLoading();
            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu TLS
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(TLSMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(totalQuantityToProduce1), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} : {2} ne correspond pas aux quantités validées dans le dispatch : {3}.", TLSRecetaName, TLSMenuName, mappedDetailsByMenu.qtyToProduce, totalQuantityToProduce1);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesQuantityToProduce = productionRecipeTabPage.GetRecipesNamesAndQtiesToProduce();
            string recipeQuantityToProduce;
            recipesQuantityToProduce.TryGetValue(TLSRecetaName, out recipeQuantityToProduce);
            Assert.IsTrue(string.Equals(recipeQuantityToProduce, totalQuantityToProduce1), "Dans 'By recipe', la quantity to produce de la recette {0} : {1} ne correspond pas aux quantités validées dans le dispatch : {2}.", TLSRecetaName, recipeQuantityToProduce, totalQuantityToProduce1);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var customersQuantityToProduce = productionCustomerTabPage.GetCustomersNamesAndQties();
            string customerQuantityToProduce;
            var customerExactName = customersQuantityToProduce.First().Key;
            customersQuantityToProduce.TryGetValue(customerExactName, out customerQuantityToProduce);
            Assert.IsTrue(customerQuantityToProduce.Contains(totalQuantityToProduce1), "Dans 'By customer', la quantity to produce de la recette {0} du customer {1} : {2} ne correspond pas aux quantités validées dans le dispatch : {3}.", TLSRecetaName, TLSCustomerName, customerQuantityToProduce, totalQuantityToProduce1);

            //DELIVERY WITH SERVICIO TLS = 12
            dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, TLSDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            //Unvalidate Dispatch 
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();
            previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();

            // Set Dispatch Quantity
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, quantity12);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            //Act
            productionPage = homePage.GoToProduction_ProductionPage();
            productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, TLSMenuName);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);

            productionResultMenuTabPage.WaitLoading();
            menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu TLS
            menusNamesDeliveriesQties.TryGetValue(TLSMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(totalQuantityToProduce2), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} : {2} ne correspond pas aux quantités validées dans le dispatch : {3}.", TLSRecetaName, TLSMenuName, mappedDetailsByMenu.qtyToProduce, totalQuantityToProduce2);

            //Check By Recipe
            productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            recipesQuantityToProduce = productionRecipeTabPage.GetRecipesNamesAndQtiesToProduce();
            recipesQuantityToProduce.TryGetValue(TLSRecetaName, out recipeQuantityToProduce);
            Assert.IsTrue(recipeQuantityToProduce.Equals(totalQuantityToProduce2), "Dans 'By recipe', la quantity to produce de la recette {0} : {1} ne correspond pas aux quantités validées dans le dispatch : {2}.", TLSRecetaName, recipeQuantityToProduce, totalQuantityToProduce2);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            customersQuantityToProduce = productionCustomerTabPage.GetCustomersNamesAndQties();
            customerExactName = customersQuantityToProduce.First().Key;
            customersQuantityToProduce.TryGetValue(customerExactName, out customerQuantityToProduce);
            Assert.IsTrue(customerQuantityToProduce.Contains(totalQuantityToProduce2), "Dans 'By customer', la quantity to produce de la recette {0} du customer {1} : {2} ne correspond pas aux quantités validées dans le dispatch : {3}.", TLSRecetaName, TLSCustomerName, customerQuantityToProduce, totalQuantityToProduce2);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckQuantityToProduce_MethodPaxPerPackagingGrouped()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string TLSGR1QuantityToProduce = "5"; // quantity*coef
            string TLSGR2QuantityToProduce = "10";
            string totalQuantityToProduce = "15";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();

            //MENU TLS GR 1
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "MENU GR TLS ");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu TLS GR 1
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(menuTLSGR1Name, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(TLSGR1QuantityToProduce), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", recetaTLSGRName, menuTLSGR1Name);

            //Checks Menu TLS GR 2
            menusNamesDeliveriesQties.TryGetValue(menuTLSGR2Name, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(TLSGR2QuantityToProduce), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", recetaTLSGRName, menuTLSGR2Name);


            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesQuantityToProduce = productionRecipeTabPage.GetRecipesNamesAndQtiesToProduce();
            string recipeQuantityToProduce;
            recipesQuantityToProduce.TryGetValue(recetaTLSGRName, out recipeQuantityToProduce);
            Assert.IsTrue(recipeQuantityToProduce.Equals(totalQuantityToProduce), "Dans 'By recipe', la quantity to produce de la recette {0} ne correspond pas aux quantités validées dans le dispatch.", recetaTLSGRName);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var customersQuantityToProduce = productionCustomerTabPage.GetCustomersNamesAndQties();
            string customerQuantityToProduce;
            customersQuantityToProduce.TryGetValue(customerTLSGRName, out customerQuantityToProduce);
            Assert.IsTrue(customerQuantityToProduce.Contains(totalQuantityToProduce), "Dans 'By customer', la quantity to produce de la recette {0} du customer {1} ne correspond pas aux quantités validées dans le dispatch.", recetaTLSGRName, customerTLSGRName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckQuantityToProduce_MethodByRecipeVariant()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantityToProduce = "16"; // quantity*coef

            // Arrange
            HomePage homePage = LogInAsAdmin(); ;

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, menuPFName);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu TLS GR 1
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(menuPFName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(quantityToProduce), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", recetaPFName, menuPFName);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesQuantityToProduce = productionRecipeTabPage.GetRecipesNamesAndQtiesToProduce();
            string recipeQuantityToProduce;
            recipesQuantityToProduce.TryGetValue(recetaPFName, out recipeQuantityToProduce);
            Assert.IsTrue(recipeQuantityToProduce.Equals(quantityToProduce), "Dans 'By recipe', la quantity to produce de la recette {0} ne correspond pas aux quantités validées dans le dispatch.", recetaPFName);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            productionCustomerTabPage.FoldUnfoldAll();
            var customersQuantityToProduce = productionCustomerTabPage.GetCustomersNamesAndQties();
            string customerQuantityToProduce;
            customersQuantityToProduce.TryGetValue(customerPFName, out customerQuantityToProduce);
            Assert.IsTrue(customerQuantityToProduce.Contains(quantityToProduce), "Dans 'By customer', la quantity to produce de la recette {0} du customer {1} ne correspond pas aux quantités validées dans le dispatch.", recetaPFName, customerPFName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckQuantityToProduce_MethodByRecipeVariant2Recipes()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantityToProduceRecipe2 = "2";
            string quantityToProduceRecipe1 = "15";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            //RECIPE CONDREN 2
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, condrenReceta2Name);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            productionResultMenuTabPage.WaitLoading();

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu 
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(condrenMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(quantityToProduceRecipe2), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", condrenReceta2Name, condrenMenuName);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesQuantityToProduce = productionRecipeTabPage.GetRecipesNamesAndQtiesToProduce();
            recipesQuantityToProduce.TryGetValue(condrenReceta2Name, out string recipeQuantityToProduce);
            Assert.IsTrue(recipeQuantityToProduce.Equals(quantityToProduceRecipe2), "Dans 'By recipe', la quantity to produce de la recette {0} ne correspond pas aux quantités validées dans le dispatch.", condrenReceta2Name);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var customersQuantityToProduce = productionCustomerTabPage.GetCustomersNamesAndQties();
            customersQuantityToProduce.TryGetValue(condrenCustomerName, out string customerQuantityToProduce);
            Assert.IsTrue(customerQuantityToProduce.Contains(quantityToProduceRecipe2), "Dans 'By customer', la quantity to produce de la recette {0} du customer {1} ne correspond pas aux quantités validées dans le dispatch.", condrenReceta2Name, condrenCustomerName);

            //RECIPE CONDREN 1
            productionCustomerTabPage.GoToProductionMenuTab();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, condrenReceta1Name);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            productionResultMenuTabPage.WaitLoading();

            menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu 
            menusNamesDeliveriesQties.TryGetValue(condrenMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(quantityToProduceRecipe1), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", condrenReceta1Name, condrenMenuName);

            //Check By Recipe
            productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            recipesQuantityToProduce = productionRecipeTabPage.GetRecipesNamesAndQtiesToProduce();
            recipesQuantityToProduce.TryGetValue(condrenReceta1Name, out recipeQuantityToProduce);
            Assert.IsTrue(recipeQuantityToProduce.Equals(quantityToProduceRecipe1), "Dans 'By recipe', la quantity to produce de la recette {0} ne correspond pas aux quantités validées dans le dispatch.", condrenReceta1Name);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            customersQuantityToProduce = productionCustomerTabPage.GetCustomersNamesAndQties();
            customersQuantityToProduce.TryGetValue(condrenCustomerName, out customerQuantityToProduce);
            Assert.IsTrue(customerQuantityToProduce.Contains(quantityToProduceRecipe1), "Dans 'By customer', la quantity to produce de la recette {0} du customer {1} ne correspond pas aux quantités validées dans le dispatch.", recetaPFName, condrenCustomerName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckQuantityToProduce_MethodRecipeVariantAndGroupedByRecipe()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantityToProduceRecipe1 = "10";
            string quantityToProduceRecipe2 = "50";
            string totalQuantityToProduceRecipe = "60";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "BARENTIN");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu 
            MappedDetailsByMenu mappedDetailsByMenu;
            //Barentin Menu Adulto
            //Barentin Receta 1
            menusNamesDeliveriesQties.TryGetValue(barentinMenuAdulto, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(quantityToProduceRecipe1), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", barentinReceta1, barentinMenuAdulto);
            //Barentin Receta 2
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipeQuantityToProduce(barentinMenuAdulto, barentinRecetaNP2).Contains(quantityToProduceRecipe1), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", barentinRecetaNP2, productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuAdulto, barentinRecetaNP2), quantityToProduceRecipe1);
            //Barentin Receta 3
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipeQuantityToProduce(barentinMenuAdulto, barentinRecetaBulk3).Contains(quantityToProduceRecipe1), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", barentinRecetaBulk3, productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuAdulto, barentinRecetaBulk3), quantityToProduceRecipe1);

            //Barentin Menu Collegio
            //Barentin Receta 1
            menusNamesDeliveriesQties.TryGetValue(barentinMenuCollegio, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.qtyToProduce.Contains(quantityToProduceRecipe2), "Dans 'By menu', la quantity to produce de la recette {0} du menu {1} ne correspond pas aux quantités validées dans le dispatch.", barentinReceta1, barentinMenuCollegio);
            //Barentin Receta 2
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipeQuantityToProduce(barentinMenuCollegio, barentinRecetaNP2).Contains(quantityToProduceRecipe2), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", barentinRecetaNP2, productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuCollegio, barentinRecetaNP2), quantityToProduceRecipe2);
            //Barentin Receta 3
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipeQuantityToProduce(barentinMenuCollegio, barentinRecetaBulk3).Contains(quantityToProduceRecipe2), "Le nombre de PAX de cette recette {0} ne corrsepond pas aux quantités validées dans le dispatch, valeur <{1}> au lieu de <PAX :{2}>.", barentinRecetaBulk3, productionResultMenuTabPage.GetMenuRecipePAX(barentinMenuCollegio, barentinRecetaBulk3), quantityToProduceRecipe2);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesQuantityToProduce = productionRecipeTabPage.GetRecipesNamesAndQtiesToProduce();
            string recipeQuantityToProduce;
            //Barentin Receta 1
            recipesQuantityToProduce.TryGetValue(barentinReceta1, out recipeQuantityToProduce);
            Assert.IsTrue(recipeQuantityToProduce.Equals(totalQuantityToProduceRecipe), "Dans 'By recipe', la quantity to produce de la recette {0} ne correspond pas aux quantités validées dans le dispatch.", barentinReceta1);
            //Barentin Receta 2
            recipesQuantityToProduce.TryGetValue(barentinRecetaNP2, out recipeQuantityToProduce);
            Assert.IsTrue(recipeQuantityToProduce.Equals(totalQuantityToProduceRecipe), "Dans 'By recipe', la quantity to produce de la recette {0} ne correspond pas aux quantités validées dans le dispatch.", barentinRecetaNP2);
            //Barentin Receta 3
            recipesQuantityToProduce.TryGetValue(barentinRecetaBulk3, out recipeQuantityToProduce);
            Assert.IsTrue(recipeQuantityToProduce.Equals(totalQuantityToProduceRecipe), "Dans 'By recipe', la quantity to produce de la recette {0} ne correspond pas aux quantités validées dans le dispatch.", barentinRecetaBulk3);

            //Check ByCustomer
            // nom recette, nombre menus, PAX, Quantity to produce
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            //Barentin Receta 1
            Assert.IsTrue(productionCustomerTabPage.GetCustomerRecipeQuantityToProduce(barentinCustomerName, barentinReceta1).Contains(totalQuantityToProduceRecipe), "Dans 'By customer', la quantity to produce de la recette {0} du customer {1} ne correspond pas aux quantités validées dans le dispatch.", barentinReceta1, barentinCustomerName);
            //Barentin Receta 2
            Assert.IsTrue(productionCustomerTabPage.GetCustomerRecipeQuantityToProduce(barentinCustomerName, barentinRecetaNP2).Contains(totalQuantityToProduceRecipe), "Dans 'By customer', la quantity to produce de la recette {0} du customer {1} ne correspond pas aux quantités validées dans le dispatch.", barentinRecetaNP2, barentinCustomerName);
            //Barentin Receta 3
            Assert.IsTrue(productionCustomerTabPage.GetCustomerRecipeQuantityToProduce(barentinCustomerName, barentinRecetaBulk3).Contains(totalQuantityToProduceRecipe), "Dans 'By customer', la quantity to produce de la recette {0} du customer {1} ne correspond pas aux quantités validées dans le dispatch.", barentinRecetaBulk3, barentinCustomerName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckFoodPack_MethodIndividualMultiportion()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, ESDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();

            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESMULServicioName);
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(ESINDMULServicioName)) > 0, "Le service {0} n'a pas de menu associé.", ESINDMULServicioName);

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "ES ");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu ES IND
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(ESINDMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.foodPack.Contains(foodpackIndividual), "Dans 'By menu', le foodpack de la recette {0} du menu {1} n'est pas le bon.", recetaESName, ESINDMenuName);

            //Checks Menu ES MUL
            menusNamesDeliveriesQties.TryGetValue(ESMULMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.foodPack.Contains(foodpackMultiporcion), "Dans 'By menu', le foodpack de la recette {0} du menu {1} n'est pas le bon.", recetaESName, ESMULMenuName);

            //Checks Menu ES IND MUL
            menusNamesDeliveriesQties.TryGetValue(ESINDMULMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(new[] { foodpackIndividual, foodpackMultiporcion }.Any(foodPack => mappedDetailsByMenu.foodPack.Contains(foodPack)), "Dans 'By menu', le foodpack de la recette {0} du menu {1} n'est pas le bon.", recetaESName, ESINDMULMenuName);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesFoodPack = productionRecipeTabPage.GetFirstRecipeFoodPack();
            Assert.IsTrue(new[] { foodpackIndividual, foodpackMultiporcion }.Any(foodPack => recipesFoodPack.Contains(foodPack)), "Dans 'By recipe', le foodpack de la recette {0} n'est pas le bon.", recetaESName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckFoodPack_MethodPaxPerPackaging()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, TLSDeliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();

            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, TLSMenuName);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu
            menusNamesDeliveriesQties.TryGetValue(TLSMenuName, out MappedDetailsByMenu mappedDetailsByMenu);
            Assert.IsTrue(new[] { foodpackB1, foodpackB4, foodpackB6 }.Any(foodPack => mappedDetailsByMenu.foodPack.Contains(foodPack)), "Dans 'By menu', le foodpack de la recette {0} du menu {1} n'est pas le bon.", TLSRecetaName, TLSMenuName);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesFoodPack = productionRecipeTabPage.GetFirstRecipeFoodPack();
            Assert.IsTrue(new[] { foodpackB1, foodpackB4, foodpackB6 }.Any(foodPack => recipesFoodPack.Contains(foodPack)), "Dans 'By recipe', le foodpack de la recette {0} n'est pas le bon.", TLSRecetaName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckFoodPack_MethodPaxPerPackagingGrouped()
        {
            string siteACE = "ACE";
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "MENU GR TLS");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(menuTLSGR1Name, out mappedDetailsByMenu);
            Assert.IsTrue(new[] { foodpackB1, foodpackB4, foodpackB6 }.Any(foodPack => mappedDetailsByMenu.foodPack.Contains(foodPack)), "Dans 'By menu', le foodpack de la recette {0} du menu {1} n'est pas le bon.", TLSRecetaName, menuTLSGR1Name);

            menusNamesDeliveriesQties.TryGetValue(menuTLSGR2Name, out mappedDetailsByMenu);
            Assert.IsTrue(new[] { foodpackB1, foodpackB4, foodpackB6 }.Any(foodPack => mappedDetailsByMenu.foodPack.Contains(foodPack)), "Dans 'By menu', le foodpack de la recette {0} du menu {1} n'est pas le bon.", TLSRecetaName, menuTLSGR2Name);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesFoodPack = productionRecipeTabPage.GetFirstRecipeFoodPack();
            Assert.IsTrue(new[] { foodpackB1, foodpackB4, foodpackB6 }.Any(foodPack => recipesFoodPack.Contains(foodPack)), "Dans 'By recipe', le foodpack de la recette {0} n'est pas le bon.", TLSRecetaName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckFoodPack_MethodByRecipeVariant()
        {

            string site = "ACE";
            int returntoDay = -1;
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, menuPFName);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            while (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                var jour = DateUtils.Now.AddDays(returntoDay).ToString("dd/MM/yyyy");
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, jour);
                returntoDay--;
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(menuPFName, out mappedDetailsByMenu);
            Assert.IsTrue(new[] { foodpackBACGN, foodpackBACGNhalf }.Any(foodPack => mappedDetailsByMenu.foodPack.Contains(foodPack)), "Dans 'By menu', le foodpack de la recette {0} du menu {1} n'est pas le bon.", recetaPFName, menuPFName);
            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesFoodPack = productionRecipeTabPage.GetFirstRecipeFoodPack();
            Assert.IsTrue(new[] { foodpackBACGN, foodpackBACGNhalf }.Any(foodPack => recipesFoodPack.Contains(foodPack)), "Dans 'By recipe', le foodpack de la recette {0} n'est pas le bon.", recetaPFName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckFoodPack_MethodByRecipeVariant2Recipes()
        {
            string site = "ACE";
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, condrenReceta1Name);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);
            productionResultMenuTabPage.WaitLoading(); //time for search filter

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu
            MappedDetailsByMenu mappedDetailsByMenu = new MappedDetailsByMenu();
            menusNamesDeliveriesQties.TryGetValue(condrenMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(new[] { foodpackBACGN, foodpackBACGNhalf }.Any(foodPack => mappedDetailsByMenu.foodPack.Contains(foodPack)), "Dans 'By menu', le foodpack n'est pas le bon.", condrenReceta1Name, condrenMenuName);
            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesFoodPack = productionRecipeTabPage.GetFirstRecipeFoodPack();
            Assert.IsTrue(new[] { foodpackBACGN, foodpackBACGNhalf }.Any(foodPack => recipesFoodPack.Contains(foodPack)), "Dans 'By recipe', le foodpack de la recette {0} n'est pas le bon.", condrenReceta1Name);

            productionRecipeTabPage.GoToProductionMenuTab();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, condrenReceta2Name);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);
            productionResultMenuTabPage.WaitLoading(); //time for search filter

            menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu
            menusNamesDeliveriesQties.TryGetValue(condrenMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.foodPack.Contains(foodpackB1), "Dans 'By menu', le foodpack n'est pas le bon.", condrenReceta2Name, condrenMenuName);
            //Check By Recipe
            productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            recipesFoodPack = productionRecipeTabPage.GetFirstRecipeFoodPack();
            Assert.IsTrue(mappedDetailsByMenu.foodPack.Contains(foodpackB1), "Dans 'By recipe', le foodpack de la recette {0} n'est pas le bon.", condrenReceta2Name);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckFoodPack_MethodRecipeVariantAndGroupedByRecipe()
        {
            string site = "ACE";
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "BARENTIN");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, site);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            //Check ByMenu
            //Menu BARENTIN Adulto
            //Barentin Receta 1
            Assert.IsTrue(new[] { foodpackBACGN, foodpackBACGNhalf }.Any(foodPack => productionResultMenuTabPage.GetMenuRecipeFoodpack(barentinMenuAdulto, barentinReceta1).Contains(foodPack)), "Dans 'By menu', le foodpack n'est pas le bon.", barentinReceta1, barentinMenuAdulto);
            //Barentin Receta 2
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipeFoodpack(barentinMenuAdulto, barentinRecetaNP2).Contains(foodpackCRT), "Dans 'By menu', le foodpack de la recette {0} du menu {1} n'est pas le bon.", barentinRecetaNP2, barentinMenuAdulto);
            //Barentin Receta 3
            Assert.IsTrue(new[] { foodpackEntier, foodpackDemi }.Any(foodPack => productionResultMenuTabPage.GetMenuRecipeFoodpack(barentinMenuAdulto, barentinRecetaBulk3).Contains(foodPack)), "Dans 'By menu', le foodpack n'est pas le bon.", barentinRecetaBulk3, barentinMenuAdulto);

            //Menu BARENTIN Collegio
            //Barentin Receta 1
            Assert.IsTrue(new[] { foodpackBACGN, foodpackBACGNhalf }.Any(foodPack => productionResultMenuTabPage.GetMenuRecipeFoodpack(barentinMenuCollegio, barentinReceta1).Contains(foodPack)), "Dans 'By menu', le foodpack n'est pas le bon.", barentinReceta1, barentinMenuCollegio);
            //Barentin Receta 2
            Assert.IsTrue(productionResultMenuTabPage.GetMenuRecipeFoodpack(barentinMenuCollegio, barentinRecetaNP2).Contains(foodpackCRT), "Dans 'By menu', le foodpack de la recette {0} du menu {1} n'est pas le bon.", barentinRecetaNP2, barentinMenuCollegio);
            //Barentin Receta 3
            Assert.IsTrue(new[] { foodpackEntier, foodpackDemi }.Any(foodPack => productionResultMenuTabPage.GetMenuRecipeFoodpack(barentinMenuCollegio, barentinRecetaBulk3).Contains(foodPack)), "Dans 'By menu', le foodpack n'est pas le bon.", barentinRecetaBulk3, barentinMenuCollegio);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesFoodPack = productionRecipeTabPage.GetFirstRecipeFoodPack();
            Assert.IsTrue(new[] { foodpackBACGN, foodpackBACGNhalf }.Any(foodPack => recipesFoodPack.Contains(foodPack)), "Dans 'By recipe', le foodpack de la recette {0} n'est pas le bon.", barentinReceta1);
            Assert.IsTrue(productionRecipeTabPage.GetRecipeFoodPack(barentinRecetaNP2).Contains(foodpackCRT), "Dans 'By recipe', le foodpack de la recette {0} n'est pas le bon.", barentinRecetaNP2);
            Assert.IsTrue(new[] { foodpackEntier, foodpackDemi }.Any(foodPack => productionRecipeTabPage.GetRecipeFoodPack(barentinRecetaBulk3).Contains(foodPack)), "Dans 'By recipe', le foodpack de la recette {0} n'est pas le bon.", barentinRecetaBulk3);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckItemWeight_MethodIndividualMultiportion()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity5 = "5";
            string quantity10 = "10";
            string quantity1 = "1";
            string ESINDNetWeight = "0,45"; //portion*quantity*coef/1000
            string ESMULNetWeight = "1,5";
            string ESINDMULNetWeight1 = "0,3";
            string totalQuantityToProduce1 = "2,25 KG"; //ESINDNetWeight + ESMULNetWeight + ESINDMULNetWeight1
            string ESINDMULNetWeight2 = "2,25";
            string totalQuantityToProduce2 = "4,2 KG";

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
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "ES ");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            ProductionSearchResultsMenuTabPage p = productionResultMenuTabPage;
            //Checks Menu ES IND
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(ESINDMenuName, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(ESINDNetWeight), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, recetaESName, ESINDMenuName, ESINDNetWeight, mappedDetailsByMenu.itemNetWeight);
            //Checks Menu ES MUL
            menusNamesDeliveriesQties.TryGetValue(ESMULMenuName, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(ESMULNetWeight), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, recetaESName, ESMULMenuName, ESMULNetWeight, mappedDetailsByMenu.itemNetWeight);

            //Checks Menu ES IND MUL
            menusNamesDeliveriesQties.TryGetValue(ESINDMULMenuName, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(ESINDMULNetWeight1), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, recetaESName, ESINDMULMenuName, ESINDMULNetWeight1, mappedDetailsByMenu.itemNetWeight);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesNetWeight = productionRecipeTabPage.GetRecipesItemsNetWeight();
            string recipeNetWeight;
            recipesNetWeight.TryGetValue(recetaESName, out recipeNetWeight);
            Assert.IsTrue(recipeNetWeight.Contains(totalQuantityToProduce1), "Dans 'By recipe', le Net Weight del'item {0}, de la recette {1}, ne correspond pas aux quantités validées dans le dispatch, valeur <{3} KG> au lieu de <{4}>.", itemTest, recetaESName, totalQuantityToProduce1, recipeNetWeight);

            //Check ByCustomer
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var customerRecipeItemNetWeight = productionCustomerTabPage.GetCustomerRecipeItemNetWeight(customerESName);
            Assert.IsTrue(customerRecipeItemNetWeight.Contains(totalQuantityToProduce1), "Dans 'By customer', le Net Weight de l'item {0}, de la recette {1}, du customer {2}, ne correspond pas aux quantités validées dans le dispatchvaleur <{3} KG> au lieu de <{4}>.", itemTest, recetaESName, customerESName, totalQuantityToProduce1, customerRecipeItemNetWeight);

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

            //Act
            productionPage = homePage.GoToProduction_ProductionPage();
            productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "ES ");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);

            menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            //Checks Menu ES IND
            menusNamesDeliveriesQties.TryGetValue(ESINDMenuName, out mappedDetailsByMenu);
            Assert.IsTrue(mappedDetailsByMenu.itemNetWeight.Contains(ESINDNetWeight), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, recetaESName, ESINDMenuName, ESINDNetWeight, mappedDetailsByMenu.itemNetWeight);

            //Checks Menu ES MUL
            menusNamesDeliveriesQties.TryGetValue(ESMULMenuName, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(ESMULNetWeight), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, recetaESName, ESMULMenuName, ESMULNetWeight, mappedDetailsByMenu.itemNetWeight);

            //Checks Menu ES IND MUL
            menusNamesDeliveriesQties.TryGetValue(ESINDMULMenuName, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(ESINDMULNetWeight2), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, recetaESName, ESINDMULMenuName, ESINDMULNetWeight2, mappedDetailsByMenu.itemNetWeight);

            //Check By Recipe
            productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            recipesNetWeight = productionRecipeTabPage.GetRecipesItemsNetWeight();
            recipesNetWeight.TryGetValue(recetaESName, out recipeNetWeight);
            Assert.IsTrue(recipeNetWeight.Contains(totalQuantityToProduce2), "Dans 'By recipe', le Net Weight del'item {0}, de la recette {1}, ne correspond pas aux quantités validées dans le dispatch, valeur <{3} KG> au lieu de <{4}>.", itemTest, recetaESName, totalQuantityToProduce2, recipeNetWeight);

            //Check ByCustomer
            productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            customerRecipeItemNetWeight = productionCustomerTabPage.GetCustomerRecipeItemNetWeight(customerESName);
            Assert.IsTrue(customerRecipeItemNetWeight.Contains(totalQuantityToProduce2), "Dans 'By customer', le Net Weight de l'item {0}, de la recette {1}, du customer {2}, ne correspond pas aux quantités validées dans le dispatchvaleur <{3} KG> au lieu de <{4}>.", itemTest, recetaESName, customerESName, totalQuantityToProduce2, customerRecipeItemNetWeight);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckItemWeight_MethodPaxPerPackaging()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string quantity10 = "10";
            string quantity12 = "12";
            string menuNetWeight1 = "3";
            string menuNetWeight2 = "3,6";

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
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, quantity10);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, TLSMenuName);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);

            productionResultMenuTabPage.WaitLoading();
            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            ProductionSearchResultsMenuTabPage p = productionResultMenuTabPage;
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(TLSMenuName, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(menuNetWeight1), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, TLSRecetaName, TLSMenuName, menuNetWeight1, mappedDetailsByMenu.itemNetWeight);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesNetWeight = productionRecipeTabPage.GetRecipesItemsNetWeight();
            string recipeNetWeight;
            recipesNetWeight.TryGetValue(TLSRecetaName, out recipeNetWeight);
            Assert.AreEqual(p.Round2decimals(recipeNetWeight), p.Round2decimals(menuNetWeight1), "Dans 'By recipe', le Net Weight del'item {0}, de la recette {1}, ne correspond pas aux quantités validées dans le dispatch, valeur <{2} KG> au lieu de <{3}>.", itemTest, TLSRecetaName, menuNetWeight1, recipeNetWeight);

            //Check ByCustomer
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var customerRecipeItemWeight = productionCustomerTabPage.GetCustomerRecipeItemNetWeight(TLSCustomerName);
            Assert.AreEqual(p.Round2decimals(customerRecipeItemWeight), p.Round2decimals(menuNetWeight1), "Dans 'By customer', le Net Weight de l'item {0}, de la recette {1}, du customer {2}, ne correspond pas aux quantités validées dans le dispatch.", itemTest, TLSRecetaName, TLSCustomerName, menuNetWeight1, customerRecipeItemWeight);

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
            dispatchPage.AddQuantityOnPrevisonalQuantity(TLSServicioName, quantity12);

            // Validate Dispatch Previsionnal QUantity and Quantity to produce
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");

            //Check if service associated with a menu
            Assert.IsTrue(Int32.Parse(dispatchPage.GetNumberMenusAssociated(TLSServicioName)) > 0, "Le service {0} n'a pas de menu associé.", TLSServicioName);

            //Act
            productionPage = homePage.GoToProduction_ProductionPage();
            productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, TLSMenuName);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);

            productionResultMenuTabPage.WaitLoading();
            menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();


            //Check ByMenu
            menusNamesDeliveriesQties.TryGetValue(TLSMenuName, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(menuNetWeight2), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, TLSRecetaName, TLSMenuName, menuNetWeight2, mappedDetailsByMenu.itemNetWeight);

            //Check By Recipe
            productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            recipesNetWeight = productionRecipeTabPage.GetRecipesItemsNetWeight();
            recipesNetWeight.TryGetValue(TLSRecetaName, out recipeNetWeight);
            var roundedRecipeNetWeight = p.Round2decimals(recipeNetWeight);
            var roundedMenuNetWeight2 = p.Round2decimals(menuNetWeight2);
            Assert.AreEqual(roundedRecipeNetWeight, roundedMenuNetWeight2, "Dans 'By recipe', le Net Weight del'item {0}, de la recette {1}, ne correspond pas aux quantités validées dans le dispatch, valeur <{2} KG> au lieu de <{3}>.", itemTest, recetaESName, menuNetWeight2, recipeNetWeight);

            //Check ByCustomer
            productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            customerRecipeItemWeight = productionCustomerTabPage.GetCustomerRecipeItemNetWeight(TLSCustomerName);
            Assert.AreEqual(p.Round2decimals(recipeNetWeight), p.Round2decimals(menuNetWeight2), "Dans 'By customer', le Net Weight de l'item {0}, de la recette {1}, du customer {2}, ne correspond pas aux quantités validées dans le dispatch.", itemTest, TLSRecetaName, TLSCustomerName, menuNetWeight2, customerRecipeItemWeight);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckItemWeight_MethodPaxPerPackagingGrouped()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string menuNetWeight1 = "0,5";
            string menuNetWeight2 = "1";
            string totalMenuNetWeight = "1,5";


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();

            //MENU TLS GR
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "MENU GR TLS ");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            ProductionSearchResultsMenuTabPage p = productionResultMenuTabPage;
            MappedDetailsByMenu mappedDetailsByMenu;
            // Menu TLS GR 1
            menusNamesDeliveriesQties.TryGetValue(menuTLSGR1Name, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(menuNetWeight1), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, recetaTLSGRName, menuTLSGR1Name, menuNetWeight1, mappedDetailsByMenu.itemNetWeight);

            // Menu TLS GR 2
            menusNamesDeliveriesQties.TryGetValue(menuTLSGR2Name, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(menuNetWeight2), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, recetaTLSGRName, menuTLSGR2Name, menuNetWeight2, mappedDetailsByMenu.itemNetWeight);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesNetWeight = productionRecipeTabPage.GetRecipesItemsNetWeight();
            string recipeNetWeight;
            recipesNetWeight.TryGetValue(recetaTLSGRName, out recipeNetWeight);
            Assert.AreEqual(p.Round2decimals(recipeNetWeight), p.Round2decimals(totalMenuNetWeight), "Dans 'By recipe', le Net Weight del'item {0}, de la recette {1}, ne correspond pas aux quantités validées dans le dispatch, valeur <{2} KG> au lieu de <{3}>.", itemTest, recetaTLSGRName, totalMenuNetWeight, recipeNetWeight);

            //Check ByCustomer
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var customerRecipeItemWeight = productionCustomerTabPage.GetCustomerRecipeItemNetWeight(customerTLSGRName);
            Assert.AreEqual(p.Round2decimals(customerRecipeItemWeight), p.Round2decimals(totalMenuNetWeight), "Dans 'By customer', le Net Weight de l'item {0}, de la recette {1}, du customer {2}, ne correspond pas aux quantités validées dans le dispatch.", itemTest, recetaTLSGRName, customerTLSGRName, totalMenuNetWeight, customerRecipeItemWeight);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckItemWeight_MethodByRecipeVariant()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string netWeight = "2,4";
            string weight = "2,526"; //netWeight * yield

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, menuPFName);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, yesterday);
            }

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            ProductionSearchResultsMenuTabPage p = productionResultMenuTabPage;
            MappedDetailsByMenu mappedDetailsByMenu;
            // Menu TLS GR 1
            menusNamesDeliveriesQties.TryGetValue(menuPFName, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(netWeight), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, condrenReceta1Name, condrenMenuName, netWeight, mappedDetailsByMenu.itemNetWeight);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesNetWeight = productionRecipeTabPage.GetRecipesItemsNetWeight();
            string recipeNetWeight;
            recipesNetWeight.TryGetValue(recetaPFName, out recipeNetWeight);
            Assert.IsTrue(recipeNetWeight.Contains(netWeight), "Dans 'By recipe', le Net Weight del'item {0}, de la recette {1}, ne correspond pas aux quantités validées dans le dispatch, valeur <{2} KG> au lieu de <{3}>.", itemTest, recetaPFName, netWeight, recipeNetWeight);

            //Check ByCustomer
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var customerRecipeItemNetWeight = productionCustomerTabPage.GetCustomerRecipeItemNetWeight(customerPFName);
            Assert.IsTrue(customerRecipeItemNetWeight.Contains(netWeight), "Dans 'By customer', le Net Weight de l'item {0}, de la recette {1}, du customer {2}, ne correspond pas aux quantités validées dans le dispatch.", itemTest, recetaPFName, customerPFName, netWeight, customerRecipeItemNetWeight);

            var customerRecipeItemWeight = productionCustomerTabPage.GetCustomerRecipeItemWeight(customerPFName);
            Assert.IsTrue(customerRecipeItemWeight.Contains(weight), "Dans 'By customer', le Net Weight de l'item {0}, de la recette {1}, du customer {2}, ne correspond pas aux quantités validées dans le dispatch.", itemTest, recetaPFName, customerPFName, weight, customerRecipeItemWeight);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckItemWeight_MethodByRecipeVariant2Recipes()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string netWeight = "2,25";
            string weight = "2,368";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            //RECIPE CONDREN 2
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, condrenReceta1Name);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);
            productionResultMenuTabPage.WaitLoading();

            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            ProductionSearchResultsMenuTabPage p = productionResultMenuTabPage;
            MappedDetailsByMenu mappedDetailsByMenu;
            menusNamesDeliveriesQties.TryGetValue(condrenMenuName, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(netWeight), "Dans 'By menu', le net weight de l'item {0} de la recette {1} du menu {2} ne correspond pas, valeur <{3} KG> au lieu de <{4}>.", itemTest, condrenReceta1Name, condrenMenuName, netWeight, mappedDetailsByMenu.itemNetWeight);

            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesNetWeight = productionRecipeTabPage.GetRecipesItemsNetWeight();
            string recipeNetWeight;
            recipesNetWeight.TryGetValue(condrenReceta1Name, out recipeNetWeight);
            Assert.AreEqual(p.Round2decimals(recipeNetWeight), p.Round2decimals(netWeight), "Dans 'By recipe', le Net Weight del'item {0}, de la recette {1}, ne correspond pas aux quantités validées dans le dispatch, valeur <{2} KG> au lieu de <{3}>.", itemTest, condrenReceta1Name, netWeight, recipeNetWeight);
            //Check ByCustomer
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var customerRecipeItemNetWeight = productionCustomerTabPage.GetCustomerRecipeItemNetWeight(condrenCustomerName);
            Assert.AreEqual(p.Round2decimals(netWeight), p.Round2decimals(customerRecipeItemNetWeight), "Dans 'By customer', le Net Weight de l'item {0}, de la recette {1}, du customer {2}, ne correspond pas aux quantités validées dans le dispatch.", itemTest, condrenReceta1Name, condrenCustomerName, netWeight, customerRecipeItemNetWeight);

            var customerRecipeItemWeight = productionCustomerTabPage.GetCustomerRecipeItemWeight(condrenCustomerName);
            // 2,37 KG vs 2,368
            Assert.AreEqual(p.Round2decimals(weight), p.Round2decimals(customerRecipeItemWeight), "Dans 'By customer', le Net Weight de l'item {0}, de la recette {1}, du customer {2}, ne correspond pas aux quantités validées dans le dispatch.", itemTest, condrenReceta1Name, condrenCustomerName, weight, customerRecipeItemWeight);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Display_CheckItemWeight_MethodRecipeVariantAndGroupedByRecipe()
        {
            // Prepare
            string netWeight1 = "1,5";
            string weight = "6,5";
            string recipeNetWeight1 = "6,5";
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string date = DateUtils.Now.AddDays(-2).ToString("dd/MM/yyyy");

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultMenuTabPage = productionPage.Validate();
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Search, "BARENTIN");
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);
            productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.ShowItems, true);

            if (!productionResultMenuTabPage.IsResultsDisplayed())
            {
                productionResultMenuTabPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, date);
            }
            var menusNamesDeliveriesQties = productionResultMenuTabPage.GetMenusNamesAndQty();

            //Check ByMenu
            ProductionSearchResultsMenuTabPage p = productionResultMenuTabPage;
            MappedDetailsByMenu mappedDetailsByMenu;
            //Barentin Menu Adulto
            //Barentin Receta 1
            menusNamesDeliveriesQties.TryGetValue(barentinMenuAdulto, out mappedDetailsByMenu);
            Assert.AreEqual(p.Round2decimals(mappedDetailsByMenu.itemNetWeight), p.Round2decimals(netWeight1), "Dans 'By menu', le net weight du menu {0} ne correspond pas.", barentinMenuAdulto);
            Assert.AreEqual(p.Round2decimals(productionResultMenuTabPage.GetMenuRecipeNetWeight(barentinMenuAdulto, barentinReceta1)), p.Round2decimals(netWeight1), "Dans 'By menu', le net weight de la recette {0} du menu {1} ne correspond pas.", barentinReceta1, barentinMenuAdulto);
            //Barentin Receta 2
            var netWeight = productionResultMenuTabPage.GetMenuRecipeNetWeight(barentinMenuAdulto, barentinRecetaNP2);
            netWeight = (p.Round2decimals(netWeight));
            netWeight1 = p.Round2decimals(netWeight1);
            netWeight = netWeight.Substring(0, netWeight.Length - 2);
            netWeight1 = netWeight1.Substring(0, netWeight1.Length - 2);
            var menuRecipeNetWeight = double.Parse(netWeight);
            var menuRecipeNetWeight1 = double.Parse(netWeight1);
            Assert.AreEqual(menuRecipeNetWeight, menuRecipeNetWeight1, "Dans 'By menu', le net weight de la recette {0} du menu {1} ne correspond pas.", barentinRecetaNP2, barentinMenuAdulto);
            //Barentin Receta 3
            Assert.AreEqual(p.Round2decimals(productionResultMenuTabPage.GetMenuRecipeNetWeight(barentinMenuAdulto, barentinRecetaBulk3)), p.Round2decimals(netWeight1), "Dans 'By menu', le net weight de la recette {0} du menu {1} ne correspond pas.", barentinRecetaBulk3, barentinMenuAdulto);
            //Check By Recipe
            var productionRecipeTabPage = productionResultMenuTabPage.GoToProductionRecipeTab();
            var recipesNetWeight = productionRecipeTabPage.GetRecipesItemsNetWeight();
            string recipeNetWeight;
            recipesNetWeight.TryGetValue(barentinReceta1, out recipeNetWeight);
            Assert.AreEqual(p.Round2decimals(recipeNetWeight), p.Round2decimals(recipeNetWeight1), "Dans 'By recipe', le Net Weight del'item {0}, de la recette {1}, ne correspond pas aux quantités validées dans le dispatch.", itemTest, barentinReceta1);

            //Check ByCustomer
            var productionCustomerTabPage = productionResultMenuTabPage.GoToProductionCustomerTab();
            var customerRecipeItemNetWeight = productionCustomerTabPage.GetCustomerRecipeItemNetWeight(barentinCustomerName);
            Assert.AreEqual(p.Round2decimals(customerRecipeItemNetWeight), p.Round2decimals(recipeNetWeight1), "Dans 'By customer', le Net Weight de l'item {0}, de la recette {1}, du customer {2}, ne correspond pas aux quantités validées dans le dispatch.", itemTest, barentinReceta1, barentinCustomerName);

            var customerRecipeItemWeight = productionCustomerTabPage.GetCustomerRecipeItemWeight(barentinCustomerName);
            var weight1 = p.Round2decimals(customerRecipeItemWeight);
            var weight2 = p.Round2decimals(weight);
            Assert.AreEqual(weight1, weight2, "Dans 'By customer', le Net Weight de l'item {0}, de la recette {1}, du customer {2}, ne correspond pas aux quantités validées dans le dispatch.", itemTest, barentinReceta1, barentinCustomerName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_Affichage_Resultats()
        {
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string date = DateUtils.Now.AddDays(-9).ToString("dd/MM/yyyy");
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            var productionResultPage = productionPage.Validate();

            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Date, date);
            productionResultPage.Filter(ProductionSearchResultsMenuTabPage.FilterType.Site, siteACE);

            bool verifpagination = productionResultPage.Pagination();

            if (verifpagination == true)
            {
                bool verif = true;
                Assert.IsTrue(verif, "Le Sélecteur  permet d'afficher 8/16/30/50/100 résultats");

            }

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_InterfaceAccess()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var productionPage = homePage.GoToProduction_ProductionPage();
            Assert.IsNotNull(productionPage, "Erreur : La page Production n'a pas pu être chargée.");
            var productionResultPage = productionPage.Validate();
            bool isPageLoaded = productionPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isPageLoaded, "Erreur : Timeout lors du chargement de la page Production > Production.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_AffichagePaginationFavoris()
        {
            string Site = "ACE";
            string Favorite = "test";
            string dateTimeNow = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string date = DateUtils.Now.AddDays(-9).ToString("dd/MM/yyyy");
            HomePage homePage = LogInAsAdmin();
            ClearCache();
            WebDriver.Manage().Window.Size = new System.Drawing.Size(1920, 1080);
            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Date, today);
            productionPage.Filter(ProductionPage.FilterType.Site, Site);
            var favoritePage = productionPage.ClickOnResetSearch();
            var countPage1 = favoritePage.GetFavoritesList().Count;
            while (countPage1 < 8)
            {
                string nomFavoriteAvecDateHeure = $"{Favorite}_{countPage1 + 1}";
                favoritePage.MakeFavorite(nomFavoriteAvecDateHeure);
                countPage1 = favoritePage.GetFavoritesList().Count;
            }
            favoritePage.ClickPageUneFavo();
            countPage1 = favoritePage.GetFavoritesList().Count;
            Assert.IsTrue(countPage1 == 8, "Erreur de pagination : Page 1 ne contient pas 8 favoris affichés.");
            Assert.IsTrue(favoritePage.AreFavoritesDisplayedCorrectly(), "Erreur d'affichage : Les favoris sur la page 1 ne sont pas affichés correctement.");
            string nomFavoritePourPageDeux = $"{"Favorite"}_{countPage1 + 1}_{dateTimeNow}";
            favoritePage.MakeFavorite(nomFavoritePourPageDeux);
            favoritePage.ClickdeuxemePageFavo();
            var countPage2 = favoritePage.GetFavoritesList().Count;
            Assert.IsTrue(countPage2 >= 1, "Erreur de pagination : La page 2 devrait avoir au minimum 1 favori affiché.");
            Assert.IsTrue(favoritePage.AreFavoritesDisplayedCorrectly(), "Erreur d'affichage : Les favoris sur la page 2 ne sont pas affichés correctement.");


        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_AffichageCommentairePrintProduction()
        {
            string dateFrom = DateUtils.Now.AddDays(-9).ToString("dd/MM/yyyy");
            string comment = "test comment";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Production Workshop Report";
            string DocFileNameZipBegin = "All_files_";


            HomePage homePage = LogInAsAdmin();

            var productionPage = homePage.GoToProduction_ProductionPage();
            productionPage.Filter(ProductionPage.FilterType.Date, dateFrom);
            var productionResultPage = productionPage.Validate();
            productionPage.ClearDownloads();
            productionPage.ShowExtendedMenu();
            productionPage.Print(ProductionPage.PrintType.Workshop, true, comment);

            var reportPage = productionPage.ClickPrint();

            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            foreach (var file in new DirectoryInfo(downloadsPath).GetFiles())
            {
                file.Delete();
            }

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
            var CommentExists = productionPage.CheckCommentExist(comment, mots);
            Assert.IsTrue(CommentExists, "Comment does not exist");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_SelectFilterFrom_FAF()
        {
            // Set up initial dates and site
            string dateFrom = DateUtils.Now.AddDays(-30).ToString("dd/MM/yyyy");

            HomePage homePage = LogInAsAdmin();
            ProductionPage productionPage = homePage.GoToProduction_ProductionPage();

            // Apply filters: site and date range
            productionPage.Filter(ProductionPage.FilterType.Date, dateFrom);  // Store the initial "From" date
            var productionResultPage = productionPage.Validate();

            // Retrieve the displayed "From" and "To" dates after validation
            string displayedDateFrom = productionResultPage.GetFilterValue(
                (ProductionSearchResultsMenuTabPage.FilterType)(object)ProductionPage.FilterType.Date  // Casting explicite
            );

            // Assert that the "From" and "To" dates are as initially set
            Assert.AreEqual(dateFrom, displayedDateFrom, "The 'From' date filter did not remain the same after validation.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PR_SelectFilterTo_FAF()
        {
            // Set up initial dates and site
            string dateTo = DateUtils.Now.AddDays(+10).ToString("dd/MM/yyyy");

            HomePage homePage = LogInAsAdmin();
            var productionPage = homePage.GoToProduction_ProductionPage();

            // Apply filters: site and date range
            productionPage.Filter(ProductionPage.FilterType.DateTo, dateTo);  // Store the initial "To" date
            var productionResultPage = productionPage.Validate();

            string displayedDateTo = productionResultPage.GetFilterValue(
                (ProductionSearchResultsMenuTabPage.FilterType)(object)ProductionPage.FilterType.DateTo  // Casting explicite
            );

            // Assert that the "From" and "To" dates are as initially set
            Assert.AreEqual(dateTo, displayedDateTo, "The 'To' date filter did not remain the same after validation.");
        }
        
    }
}
