using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.FoodCost;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;

namespace Newrest.Winrest.FunctionalTests.Menus
{
    [TestClass]
    public class FoodCostsTests : TestBase
    {
        private const int _timeout = 600000;
        // données tests
        readonly string itemFoodCostA = "Item FoodCost A";
        readonly string itemFoodCostB = "Item FoodCost B";

        //Data FoodCost A
        readonly string customerFCAName = "Customer FoodCost A";
        readonly string customerESCode = "FCA";
        readonly string ServiceFoodCostA = "Service FoodCost A";
        readonly string MenuFoodCostAName = "Menu FoodCost A";
        readonly string MenuFoodCostACoef = "100";
        readonly string recipeFoodCostAName = "Recipe FoodCost A";
        readonly string portion15g = "15";

        // Datas FoodCost B
        readonly string customerFCBName = "Customer FoodCost B";
        readonly string customerFCBCode = "FCB";
        readonly string serviceFCB = "Service FoodCost B";
        readonly string recipeFCBName = "Recipe FoodCost B";
        readonly string menuFCBName = "Menu FoodCost B";

        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void RE_FOCO_Create_Data_FoodCostA()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string group1 = TestContext.Properties["Production_ItemGroup1"].ToString();
            string workshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string supplier1 = TestContext.Properties["Production_ItemSupplier1"].ToString();
            string taxType = TestContext.Properties["Production_ItemTaxType"].ToString();
            string prodUnit = TestContext.Properties["Production_ItemProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();

            // Arrange
           
            var homePage = LogInAsAdmin();

            //Create item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemFoodCostA);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemFoodCostA, group1, workshop1, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier1);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemFoodCostA);
            }
            var FirstItemName = itemPage.GetFirstItemName();
            Assert.AreEqual(itemFoodCostA, FirstItemName, "L'item " + itemFoodCostA + " n'existe pas.");

            // Act
            //Create Customer FoodCostA
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerFCAName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerFCAName, customerESCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerFCAName);
            }
            var isContaining = customersPage.GetFirstCustomerName().Contains(customerFCAName);
            Assert.IsTrue(isContaining, "Le customer n'a pas été créé.");

            //Create Service FoodCostA
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, ServiceFoodCostA);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(ServiceFoodCostA, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerFCAName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerFCAName);
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
            servicePage.Filter(ServicePage.FilterType.Search, ServiceFoodCostA);
            isContaining = servicePage.GetFirstServiceName().Contains(ServiceFoodCostA);
            Assert.IsTrue(isContaining, "Le service " + ServiceFoodCostA + " n'existe pas.");

            //Create Recette FoodCostA
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeFoodCostAName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                recipesPage.Filter(RecipesPage.FilterType.ShowWithoutVariants, true);
                if (recipesPage.CheckTotalNumber() == 0)
                {
                    var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                    var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeFoodCostAName, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                    // 1 variante pour ACE
                    recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                    var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                    recipeVariantPage.AddIngredient(itemFoodCostA);
                    recipeVariantPage.SetTotalPortion(portion15g);

                    recipesPage = recipeVariantPage.BackToList();
                }

                else
                {
                    var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                    recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                    var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                    recipeVariantPage.AddIngredient(itemFoodCostA);
                    recipeVariantPage.SetTotalPortion(portion15g);

                    recipesPage = recipeVariantPage.BackToList();

                }
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemFoodCostA);
                }
                recipeVariantPage.SetTotalPortion(portion15g);

                recipesPage = recipeVariantPage.BackToList();
            }
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeFoodCostAName);
            var FirstRecipeName = recipesPage.GetFirstRecipeName();
            Assert.AreEqual(recipeFoodCostAName, FirstRecipeName, "La recette " + recipeFoodCostAName + " n'existe pas.");

            //Create Menu FoodCost A
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, MenuFoodCostAName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(ServiceFoodCostA);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(MenuFoodCostAName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recipeFoodCostAName);
                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeCoef(MenuFoodCostACoef);
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipeFoodCostAName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(MenuFoodCostACoef);
                }

                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipeFoodCostAName);
                    menuDayViewPage.ClickOnFirstRecipe();
                    menuDayViewPage.SetRecipeCoef(MenuFoodCostACoef);
                }

                menusPage = menuDayViewPage.BackToList();
            }
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, MenuFoodCostAName);
            var FirstMenuName = menusPage.GetFirstMenuName();
            Assert.AreEqual(MenuFoodCostAName, FirstMenuName, "Le menu n'a pas été créé.");
        }

        [TestMethod]
        [Priority(1)]
        [Timeout(_timeout)]
        public void RE_FOCO_Create_Data_FoodCostB()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();
            string group1 = TestContext.Properties["Production_ItemGroup1"].ToString();
            string workshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string supplier1 = TestContext.Properties["Production_ItemSupplier1"].ToString();
            string taxType = TestContext.Properties["Production_ItemTaxType"].ToString();
            string prodUnit = TestContext.Properties["Production_ItemProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string recipeVariantForACE = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();
            string recipeType1 = TestContext.Properties["Production_RecipeType1"].ToString();
            string recipeCookingMode1 = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeWorkshop1 = TestContext.Properties["Production_Workshop1"].ToString();
            string recipePortion1 = "15";
            string menuVariantForACE = TestContext.Properties["Production_MenuVariant1ForACE"].ToString();

            // Arrange
            
            var homePage = LogInAsAdmin();

            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemFoodCostB);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemFoodCostB, group1, workshop1, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.setYield("100");
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier1);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemFoodCostB);
            }

            // Act
            //Create Customer FoodCostB
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerFCBName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerFCBName, customerFCBCode, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerFCBName);
            }

            var isContaining = customersPage.GetFirstCustomerName().Equals(customerFCBName);
            Assert.IsTrue(isContaining, "Le customer n'a pas été créé.");

            //Create Service FoodCostB
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceFCB);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceFCB, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerFCBName, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerFCBName);
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
            servicePage.Filter(ServicePage.FilterType.Search, serviceFCB);
            isContaining = servicePage.GetFirstServiceName().Contains(serviceFCB);
            Assert.IsTrue(isContaining, "Le service " + serviceFCB + " n'existe pas.");

            //Create Receta FoodCostB
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeFCBName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                recipesPage.Filter(RecipesPage.FilterType.ShowWithoutVariants, true);
                if (recipesPage.CheckTotalNumber() == 0)
                {
                    var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                    var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeFCBName, recipeType1, recipePortion1, true, null, recipeCookingMode1, recipeWorkshop1);

                    // 1 variante pour ACE
                    recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                    var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                    recipeVariantPage.AddIngredient(itemFoodCostB);
                    recipeVariantPage.SetTotalPortion(portion15g);
                    //inutile
                    //recipeVariantPage.ClickOnFirstIngredient();

                    recipesPage = recipeVariantPage.BackToList();
                }
                else
                {
                    var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                    recipeGeneralInfosPage.AddVariantWithSite(siteACE, recipeVariantForACE);
                    var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                    recipeVariantPage.AddIngredient(itemFoodCostB);
                    recipeVariantPage.SetTotalPortion(portion15g);

                    recipesPage = recipeVariantPage.BackToList();
                }


            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();

                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(itemFoodCostB);
                }
                recipeVariantPage.SetTotalPortion(portion15g);
                //inutile
                //recipeVariantPage.ClickOnFirstIngredient();
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeFCBName);
            isContaining = recipesPage.GetFirstRecipeName().Contains(recipeFCBName);
            Assert.IsTrue(isContaining, "La recette " + recipeFCBName + " n'existe pas.");

            //Create menu FoodCostB
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuFCBName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menusCreateModalPage.AddSite(siteACE);
                menusCreateModalPage.AddService(serviceFCB);
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuFCBName, DateUtils.Now, DateUtils.Now.AddYears(3), siteACE, menuVariantForACE);
                menuDayViewPage.AddRecipe(recipeFCBName);
                menuDayViewPage.ClickOnFirstRecipe();
                menusPage = menuDayViewPage.BackToList();
            }
            else
            {
                var menuDayViewPage = menusPage.SelectFirstMenu();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipeFCBName);
                }
                menuDayViewPage.ClickOnDayAfter();
                if (!menuDayViewPage.IsRecipeDisplayed())
                {
                    menuDayViewPage.AddRecipe(recipeFCBName);
                }

                menusPage = menuDayViewPage.BackToList();
            }

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuFCBName);
            isContaining = menusPage.GetFirstMenuName().Contains(menuFCBName);
            Assert.IsTrue(isContaining, "Le menu {0} n'a pas été créé.", menuFCBName);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_FOCO_Filter_Customer()
        {
            //Prepare
            string customer = "A.F.A. LANZAROTE ";
            //Arrange
            
            //Act
            var homePage = LogInAsAdmin();
            var foodCostPage = homePage.GoToMenus_FoodCost();
            foodCostPage.Filter(FoodCostPage.FilterType.Customer, customer);
            var checkFilter = foodCostPage.CheckCustomerFilter(customer);
            //Assert
            Assert.IsTrue(checkFilter, "Filter par customer error");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_FOCO_Filter_From()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var foodCostPage = homePage.GoToMenus_FoodCost();
            DateTime from = DateTime.Now;
            foodCostPage.Filter(FoodCostPage.FilterType.From, from);
            var checkFilter = foodCostPage.CheckFromFilter(from);
            Assert.IsTrue(checkFilter, "filter from date in header value error");
            var firstDate = foodCostPage.GetFirstDate();
            var firstMondayInWeek = foodCostPage.GetMondayOfWeekDate(from);
            //assert
            Assert.IsTrue(foodCostPage.CompareDates(firstMondayInWeek, firstDate), "first date value error");

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_FOCO_Filter_To()
        {
            //arrange

            HomePage homePage = LogInAsAdmin();

            var foodCostPage = homePage.GoToMenus_FoodCost();
            DateTime to = DateTime.Now;
            foodCostPage.WaitPageLoading();
            foodCostPage.Filter(FoodCostPage.FilterType.To, to);
            var checkFilter = foodCostPage.CheckToFilter(to);
            Assert.IsTrue(checkFilter, "filter to date in header value error");
            var lastDate = foodCostPage.GetLastDate();
            var nextSunday = foodCostPage.GetNextSundayDate(to);
            var isDateCompared = foodCostPage.CompareDates(nextSunday, lastDate);
            Assert.IsTrue(isDateCompared, "first date value error");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_FOCO_Filter_Site()
        {
            //Prepare
            var menus = new List<string>();
            var site = "ACE";
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var foodCostsPage = homePage.GoToMenus_FoodCost();
            foodCostsPage.Filter(FoodCostPage.FilterType.Site, site);
            var menusNames = foodCostsPage.GetMenusNames();
            foreach (var menuName in menusNames)
            {
                if (!menus.Contains(menuName))
                {
                    menus.Add(menuName);
                }
            }
            var menusPage = homePage.GoToMenus_Menus();
            foreach (var menuName in menus)
            {
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
                var menusDayViewPage = menusPage.ClickFirstMenu();
                var siteFromMenu = menusDayViewPage.GetSite();
                //Assert
                Assert.IsTrue(site == siteFromMenu, "erreur de filtrage par site");
                menusPage = homePage.GoToMenus_Menus();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_FOCO_Edit()
        {
            // Prepare
            // Arrange
            
            var homePage = LogInAsAdmin();
            FoodCostPage foodCostsPage = homePage.GoToMenus_FoodCost();
            foodCostsPage.Filter(FoodCostPage.FilterType.Site, "ACE");
            // Act
            //"1 .Cliquer sur ˃
            foodCostsPage.Uncollapse(0);
            string menuName = foodCostsPage.GetMenusNames()[0];
            //2.Cliquer sur le stylo
            MenusWeekViewPage menu = foodCostsPage.EditPencil(menuName);
            //3.Une nouvelle page s'ouvre avec le nom du Menu"
            string menuNameWeek = menu.GetMenuName();
            //Vérifier le nom du Menu dans la nouvelle page
            Assert.AreEqual(menuName.ToUpper(), menuNameWeek, "mauvais nom de menu");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_FOCO_Fold_Unfold_List()
        {
            // Prepare
            // Arrange
            
            var homePage = LogInAsAdmin();
            FoodCostPage foodCostsPage = homePage.GoToMenus_FoodCost();
            // Act
            foodCostsPage.ValidateFilters();

            foodCostsPage.Uncollapse(0);
            string firstMenuName = foodCostsPage.GetMenusNames()[0];
            Assert.IsTrue(foodCostsPage.IsMenuNameVisible(firstMenuName), "unfold");
            foodCostsPage.Uncollapse(0);
            Assert.IsFalse(foodCostsPage.IsMenuNameVisible(firstMenuName), "fold");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_FOCO_Filter_GuestTypes()
        {
            //Prepare
            var menus = new List<string>();
            var guestType = "BASAL ADULTO/MAYORES";
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var foodCostsPage = homePage.GoToMenus_FoodCost();
            foodCostsPage.Filter(FoodCostPage.FilterType.GuestType, guestType);
            var verifyHeader = foodCostsPage.GuestTypeExistInHeaderFilter(guestType);
            Assert.IsTrue(verifyHeader, "guest type filter n'existe pas dans le header");

            var menusNames = foodCostsPage.GetMenusNames();
            foreach (var menuName in menusNames)
            {
                if (!menus.Contains(menuName))
                {
                    menus.Add(menuName);
                }
            }
            var menusPage = homePage.GoToMenus_Menus();
            foreach (var menuName in menus)
            {
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
                var menusDayViewPage = menusPage.ClickFirstMenu();
                var guestTypeExistInMenu = menusDayViewPage.IsGuestTypeExist(guestType);
                //assert
                Assert.IsTrue(guestTypeExistInMenu, "filtrage par guest type failed");
                menusPage = homePage.GoToMenus_Menus();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_FOCO_Filter_MealTypes()
        {
            //prepare
            var mealType = "ALMUERZO";
            //arrange
            HomePage homePage = LogInAsAdmin();
            //act
            var foodCostsPage = homePage.GoToMenus_FoodCost();
            foodCostsPage.Filter(FoodCostPage.FilterType.MealType, mealType);
            var verifyExist = foodCostsPage.MealTypeExistFilter(mealType);
            //assert
            Assert.IsTrue(verifyExist, "meal type filter n'existe pas dans le menu");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_FOCO_AffichageMenuCout()
        {
            //arrange
            HomePage homePage = LogInAsAdmin();
            var foodCostPage = homePage.GoToMenus_FoodCost();

            DateTime dateFrom = DateUtils.Now.AddDays(-7);
            DateTime dateTo = DateUtils.Now.AddDays(+15);

            foodCostPage.FilterFromToDate(FoodCostPage.FilterType.From, dateFrom);
            foodCostPage.FilterFromToDate(FoodCostPage.FilterType.To, dateTo);
            foodCostPage.WaitPageLoading();
            foodCostPage.Validate();
            bool isVerifiedMenuAndCout = foodCostPage.VerifiedMenuAndCout();
            //Assert
            Assert.IsTrue(isVerifiedMenuAndCout, "les menus et les coûts associés devraient apparaitre en dessous.");

        }
    }
}