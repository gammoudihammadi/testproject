using DocumentFormat.OpenXml.VariantTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Policy;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Newrest.Winrest.FunctionalTests.Menus
{
    [TestClass]
    public class MenusTests : TestBase
    {

        private const int _timeout = 600000;
        private const string MENUS_EXCEL_SHEET_NAME = "Recipes";
        private readonly string menuNameToday = "Menu-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private readonly string recipeNameToday = "Rec-" + DateUtils.Now.ToString("dd/MM/yyyy");

        [Priority(0)]
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_CreateVariant()
        {
            // Prepare
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string variant2 = TestContext.Properties["MenuVariantACE1"].ToString();
            string variant3 = TestContext.Properties["MenuVariantACE3"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string site2 = TestContext.Properties["SiteLP"].ToString();

            string guestName = variant.Substring(0, variant.IndexOf("-")).Trim();
            string guestName2 = variant2.Substring(0, variant2.IndexOf("-")).Trim();
            string guestName3 = variant3.Substring(0, variant3.IndexOf("-")).Trim();

            string meal = variant.Substring(variant.IndexOf("-") + 1).Trim();
            string meal2 = variant2.Substring(variant2.IndexOf("-") + 1).Trim();
            string meal3 = variant3.Substring(variant3.IndexOf("-") + 1).Trim();

            //Arrange
            var homePage = LogInAsAdmin();
          
            //act
            ClearCache();

            var productionPage = homePage.GoToParameters_ProductionPage();

            productionPage.GoToTab_Guest();

            if (!productionPage.IsGuestPresent(guestName))
            {
                productionPage.AddNewGuest(guestName, false, "1000");

            }

            if (!productionPage.IsGuestPresent(guestName2))
            {
                productionPage.AddNewGuest(guestName2, false, "1001");

            }

            if (!productionPage.IsGuestPresent(guestName3))
            {
                productionPage.AddNewGuest(guestName3, false, "1002");

            }

            WebDriver.Navigate().Refresh();
            productionPage.GoToTab_Guest();
            Assert.IsTrue(productionPage.IsGuestPresent(guestName), "Le guest " + guestName + "n'est pas présent.");
            Assert.IsTrue(productionPage.IsGuestPresent(guestName2), "Le guest " + guestName2 + "n'est pas présent.");
            Assert.IsTrue(productionPage.IsGuestPresent(guestName3), "Le guest " + guestName3 + "n'est pas présent.");

            productionPage.GoToTab_Variant();
            productionPage.FilterSite(site);
            productionPage.FilterGuestType(guestName);
            productionPage.FilterMealType(meal);

            if (!productionPage.IsVariantPresent(guestName, meal))
            {
                productionPage.AddNewVariant(guestName, meal, site);
            }

            productionPage.FilterSite(site2);
            productionPage.FilterGuestType(guestName2);
            productionPage.FilterMealType(meal2);

            if (!productionPage.IsVariantPresent(guestName2, meal2))
            {
                productionPage.AddNewVariant(guestName2, meal2, site2);
            }

            productionPage.FilterSite(site2);
            productionPage.FilterGuestType(guestName3);
            productionPage.FilterMealType(meal3);

            if (!productionPage.IsVariantPresent(guestName3, meal3))
            {
                productionPage.AddNewVariant(guestName3, meal3, site2);
            }

            WebDriver.Navigate().Refresh();
            productionPage.GoToTab_Variant();

            productionPage.FilterSite(site);
            productionPage.FilterGuestType(guestName);
            productionPage.FilterMealType(meal);
            //Assert 
            Assert.IsTrue(productionPage.IsVariantPresent(guestName, meal), "La variante " + variant + "n'est pas présente.");

            productionPage.FilterSite(site2);
            productionPage.FilterGuestType(guestName2);
            productionPage.FilterMealType(meal2);
            //Assert 
            Assert.IsTrue(productionPage.IsVariantPresent(guestName2, meal2), "La variante " + variant2 + "n'est pas présente.");

            productionPage.FilterSite(site2);
            productionPage.FilterGuestType(guestName3);
            productionPage.FilterMealType(meal3);
            //Assert 
            Assert.IsTrue(productionPage.IsVariantPresent(guestName3, meal3), "La variante " + variant3 + "n'est pas présente.");
        }

        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_CreateNewRecipeForMenus()
        {
            string recipeName = "RecipeForTestMenu";
            string recipeNameBis = "Recipe2ForTestMenu";
            string site = TestContext.Properties["Site"].ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string recipeSalePrice = "5";
            string recipePoints = "1";
            string TLSEnergyKCal = "600";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");

                recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipeVariantPage.SetSalesPrice(recipeSalePrice);
                recipeVariantPage.SetPointsPrice(recipePoints);
                recipeVariantPage.ClickOnDieteticTab();
                if (!recipeVariantPage.GetEnergyKCal().Equals(TLSEnergyKCal))
                {
                    recipeVariantPage.SetEnergyKCal(TLSEnergyKCal);
                }
                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInformationPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInformationPage.ClickOnFirstVariant();
                recipeVariantPage.SetSalesPrice(recipeSalePrice);
                recipeVariantPage.SetPointsPrice(recipePoints);
                recipeVariantPage.ClickOnDieteticTab();
                if (!recipeVariantPage.GetEnergyKCal().Equals(TLSEnergyKCal))
                {
                    recipeVariantPage.SetEnergyKCal(TLSEnergyKCal);
                }
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeNameBis);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameBis, recipeType, "5");

                recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipeVariantPage.SetSalesPrice(recipeSalePrice);
                recipeVariantPage.SetPointsPrice(recipePoints);
                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInformationPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInformationPage.ClickOnFirstVariant();
                recipeVariantPage.SetSalesPrice(recipeSalePrice);
                recipeVariantPage.SetPointsPrice(recipePoints);
                recipesPage = recipeVariantPage.BackToList();
            }

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            Assert.AreEqual(recipeName, recipesPage.GetFirstRecipeName(), "La recette " + recipeName + " n'a pas été créée.");

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeNameBis);
            var firstRecipeName= recipesPage.GetFirstRecipeName();
            Assert.AreEqual(recipeNameBis, firstRecipeName, "La recette " + recipeNameBis + " n'a pas été créée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_CreateNewMenu()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant);
            menusPage = menuDayViewPage.BackToList();

            //Assert
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            menusPage.UnselectServicesFromMenu();
            Assert.AreEqual(menuName, menusPage.GetFirstMenuName(), "Le menu n'a pas été créé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_Search_Menu()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            MenusPage menusPage = homePage.GoToMenus_Menus();
            MenusCreateModalPage menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menusPage = menuDayViewPage.BackToList();

            //Assert
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            string menuNameFrom = menusPage.GetFirstMenuName();
            Assert.AreEqual(menuName, menuNameFrom, string.Format(MessageErreur.FILTRE_ERRONE, "Search"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_SortBy_Id()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            
            var homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            if (menusPage.CheckTotalNumber() < 20)
            {
                AddServiceForMenu(homePage, serviceName, guestType);
                menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
            }

            if (!menusPage.isPageSizeEqualsTo100())
            {
                menusPage.PageSize("8");
                menusPage.PageSize("100");
            }

            menusPage.Filter(MenusPage.FilterType.SortBy, "Id");
            var isSortedById = menusPage.IsSortedById();

            //Assert            
            Assert.IsTrue(isSortedById, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by Id"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_SortBy_StartDate()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            
            var homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            if (menusPage.CheckTotalNumber() < 20)
            {
                AddServiceForMenu(homePage, serviceName, guestType);
                menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
            }

            if (!menusPage.isPageSizeEqualsTo100())
            {
                menusPage.PageSize("8");
                menusPage.PageSize("100");
            }

            menusPage.Filter(MenusPage.FilterType.SortBy, "StartDate");
            var isSortedByStartDate = menusPage.IsSortedByStartDate();

            //Assert            
            Assert.IsTrue(isSortedByStartDate, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by Start Date"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_SortBy_EndDate()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            if (menusPage.CheckTotalNumber() < 20)
            {
                AddServiceForMenu(homePage, serviceName, guestType);
                menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
            }

            if (!menusPage.isPageSizeEqualsTo100())
            {
                menusPage.PageSize("8");
                menusPage.PageSize("100");
            }

            menusPage.Filter(MenusPage.FilterType.SortBy, "EndDate");
            var isSortedByEndDate = menusPage.IsSortedByEndDate();

            //Assert            
            Assert.IsTrue(isSortedByEndDate, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by End Date"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_ShowActive()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.ShowActive, true);

            if (menusPage.CheckTotalNumber() < 20)
            {
                AddServiceForMenu(homePage, serviceName, guestType);
                menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.ShowActive, true);
            }

            if (!menusPage.isPageSizeEqualsTo100())
            {
                menusPage.PageSize("8");
                menusPage.PageSize("100");
            }

            //Assert            
            Assert.IsTrue(menusPage.CheckStatus(true), String.Format(MessageErreur.FILTRE_ERRONE, "Show only active"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_ShowInactive()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.ShowInactive, true);
            menusPage.UnselectServicesFromMenu();

            if (menusPage.CheckTotalNumber() < 20)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, null, false);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.ShowInactive, true);
                menusPage.UnselectServicesFromMenu();
            }

            if (!menusPage.isPageSizeEqualsTo100())
            {
                menusPage.PageSize("8");
                menusPage.PageSize("100");
            }

            //Assert            
            Assert.IsFalse(menusPage.CheckStatus(false), String.Format(MessageErreur.FILTRE_ERRONE, "Show only inactive"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_ShowAll()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();

            //Arrange
            
            var homePage = LogInAsAdmin();
            
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.ShowAll, true);
            menusPage.UnselectServicesFromMenu();

            if (menusPage.CheckTotalNumber() < 20)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, null, false);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.ShowAll, true);
                menusPage.UnselectServicesFromMenu();
            }

            var nbTotal = menusPage.CheckTotalNumber();

            WebDriver.Navigate().Refresh();
            menusPage.Filter(MenusPage.FilterType.ShowActive, true);
            menusPage.UnselectServicesFromMenu();
            var nbActive = menusPage.CheckTotalNumber();

            WebDriver.Navigate().Refresh();
            menusPage.Filter(MenusPage.FilterType.ShowInactive, true);
            menusPage.UnselectServicesFromMenu();
            var nbInactive = menusPage.CheckTotalNumber();

            //Assert            
            Assert.AreEqual(nbTotal, nbActive + nbInactive, String.Format(MessageErreur.FILTRE_ERRONE, "Show all"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_ShowAllExceptAuto()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            MenusDayViewPage menuDayViewPage;

            //Arrange
            var homePage = LogInAsAdmin();

            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.ShowAllExceptAuto, true);

            if (menusPage.CheckTotalNumber() < 20)
            {
                AddServiceForMenu(homePage, serviceName, guestType);
                menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.ShowAllExceptAuto, true);
            }
            else
            {
                menuName = menusPage.GetFirstMenuName();
            }

            menuDayViewPage = menusPage.SelectFirstMenu();
            var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
            Assert.AreEqual(menuGeneralInformationPage.GetCommercialName(), menuName, "Le menu sélectionné n'est pas celui attendu");
            Assert.IsFalse(menuGeneralInformationPage.IsClinic(), "Le menu est associé à 'Clinic' alors que le filtre ne doit pas afficher ce type de menu");

            menuGeneralInformationPage.SetClinic(true);
            menusPage = menuGeneralInformationPage.BackToList();

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.ShowAllExceptAuto, true);
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);

            Assert.AreEqual(0, menusPage.CheckTotalNumber(), string.Format(MessageErreur.FILTRE_ERRONE, "Show all except auto"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_ShowOnlyAuto()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = $"{ menuNameToday } - {new Random().Next()}";
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            MenusDayViewPage menuDayViewPage;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            MenusPage menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            AddServiceForMenu(homePage, serviceName, guestType);
            menusPage = homePage.GoToMenus_Menus();
            MenusCreateModalPage menusCreateModalPage = menusPage.MenuCreatePage();
            menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menusPage = menuDayViewPage.BackToList();

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            menusPage.Filter(MenusPage.FilterType.ShowOnlyAuto, true);

            Assert.AreEqual(0, menusPage.CheckTotalNumber(), "Le menu créé apparaît dans la liste des menus auto alors qu'il ne devrait pas.");

            menusPage.Filter(MenusPage.FilterType.ShowAllAuto, true);
            menuDayViewPage = menusPage.SelectFirstMenu();
            MenusGeneralInformationPage menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
            string nameMenu = menuGeneralInformationPage.GetCommercialName();
            Assert.AreEqual(nameMenu, menuName, "Le menu sélectionné n'est pas celui attendu");
            bool isClinic = menuGeneralInformationPage.IsClinic();
            Assert.IsFalse(isClinic, "Le menu est déjà associé à 'Clinic'.");

            menuGeneralInformationPage.SetClinic(true);
            menusPage = menuGeneralInformationPage.BackToList();

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.ShowOnlyAuto, true);
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            int count = menusPage.CheckTotalNumber();
            Assert.AreEqual(1, count, string.Format(MessageErreur.FILTRE_ERRONE, "Show only auto"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_ShowAllAuto()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            MenusDayViewPage menuDayViewPage;

            //Arrange
            
            var homePage = LogInAsAdmin();

            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.ShowAllAuto, true);

            if (menusPage.CheckTotalNumber() < 20)
            {
                AddServiceForMenu(homePage, serviceName, guestType);
                menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.ShowAllAuto, true);
            }

            int nbTotal = menusPage.CheckTotalNumber();

            menusPage.Filter(MenusPage.FilterType.ShowAllExceptAuto, true);
            int nbExceptAuto = menusPage.CheckTotalNumber();

            menusPage.Filter(MenusPage.FilterType.ShowOnlyAuto, true);
            int nbAuto = menusPage.CheckTotalNumber();

            Assert.AreEqual(nbTotal, nbExceptAuto + nbAuto, string.Format(MessageErreur.FILTRE_ERRONE, "Show all"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_DateFrom()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            DateTime fromDate = DateUtils.Now.AddYears(1);
            DateTime fromDate1 = DateUtils.Now.AddYears(5);
            DateTime fromDate2 = DateUtils.Now.AddYears(-5);
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.DateFrom, fromDate);
            menusPage.ClearDateToFilter();

            if (menusPage.CheckTotalNumber() < 20)
            {
                AddServiceForMenu(homePage, serviceName, guestType);
                menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now.AddYears(1), DateUtils.Now.AddYears(2), site, variant, serviceName);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.DateFrom, fromDate);
                menusPage.ClearDateToFilter();
            }

            if (!menusPage.isPageSizeEqualsTo100())
            {
                menusPage.PageSize("8");
                menusPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(menusPage.IsDateRespected(fromDate, default), String.Format(MessageErreur.FILTRE_ERRONE, "From"));
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.DateFrom, fromDate1);
            Assert.IsTrue(menusPage.IsDateRespected(fromDate1, default), String.Format(MessageErreur.FILTRE_ERRONE, "From"));
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.DateFrom, fromDate2);
            Assert.IsTrue(menusPage.IsDateRespected(fromDate2, default), String.Format(MessageErreur.FILTRE_ERRONE, "From"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_DateTo()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            DateTime toDate = DateUtils.Now.AddYears(1);
            DateTime toDate1 = DateUtils.Now.AddYears(5);
            DateTime toDate2 = DateUtils.Now.AddYears(-5);
            //Arrange          
            HomePage homePage = LogInAsAdmin();             
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.ClearDateFromFilter();
            menusPage.Filter(MenusPage.FilterType.DateTo, toDate);

            if (menusPage.CheckTotalNumber() < 20)
            {
                AddServiceForMenu(homePage, serviceName, guestType);
                menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now.AddYears(1), DateUtils.Now.AddYears(2), site, variant, serviceName);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.ClearDateFromFilter();
                menusPage.Filter(MenusPage.FilterType.DateTo, toDate);

            }

            if (!menusPage.isPageSizeEqualsTo100())
            {
                menusPage.PageSize("8");
                menusPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(menusPage.IsDateRespected(default, toDate), String.Format(MessageErreur.FILTRE_ERRONE, "To"));
            menusPage.ResetFilter();
            menusPage.ClearDateFromFilter();
            menusPage.Filter(MenusPage.FilterType.DateTo, toDate1);
            Assert.IsTrue(menusPage.IsDateRespected(default, toDate1), String.Format(MessageErreur.FILTRE_ERRONE, "To"));
            menusPage.ResetFilter();
            menusPage.ClearDateFromFilter();
            menusPage.Filter(MenusPage.FilterType.DateTo, toDate2);
            Assert.IsTrue(menusPage.IsDateRespected(default, toDate2), String.Format(MessageErreur.FILTRE_ERRONE, "To"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_Site()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.Site, site);

            if (menusPage.CheckTotalNumber() < 20)
            {
                AddServiceForMenu(homePage, serviceName, guestType);
                menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.Site, site);
            }

            if (!menusPage.isPageSizeEqualsTo100())
            {
                menusPage.PageSize("8");
                menusPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(menusPage.VerifySite(site), String.Format(MessageErreur.FILTRE_ERRONE, "Sites"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_Services()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            Random rnd = new Random();
            string serviceName = "Service-" + DateUtils.Now.ToShortDateString() + "-" + rnd.Next().ToString();
            string menuName = menuNameToday + "-" + rnd.Next().ToString();
            string menuName1 = menuNameToday + "-" + rnd.Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            string otherServiceName = "Service-" + DateUtils.Now.ToShortDateString() + "-" + rnd.Next().ToString();

            //Arrange        
            HomePage homePage = LogInAsAdmin();
            //On ajoute 2 services
            AddServiceForMenu(homePage, serviceName, guestType);
            AddServiceForMenu(homePage, otherServiceName, guestType);

            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menusPage = menuDayViewPage.BackToList();

            menusCreateModalPage = menusPage.MenuCreatePage();
            menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName1, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, otherServiceName);
            menusPage = menuDayViewPage.BackToList();

            menusPage.ResetFilter();   
            menusPage.Filter(MenusPage.FilterType.Service, serviceName);
            var totalNumber = menusPage.CheckTotalNumber();
            var getMenuName = menusPage.GetFirstMenuName();
            Assert.AreEqual(1, totalNumber, "Aucun menu n'est associé au service sélectionné.");
            Assert.AreEqual(getMenuName, menuName, String.Format(MessageErreur.FILTRE_ERRONE, "Services"));
            WebDriver.Navigate().Refresh();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.Service, otherServiceName);
            totalNumber = menusPage.CheckTotalNumber();
            getMenuName = menusPage.GetFirstMenuName();
            Assert.AreEqual(1, totalNumber, String.Format(MessageErreur.FILTRE_ERRONE, "Services"));
            Assert.AreEqual(getMenuName, menuName1, String.Format(MessageErreur.FILTRE_ERRONE, "Services"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_TypeGuest()
        {
            //Prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string variant1 = TestContext.Properties["MenuVariantACE1"].ToString();
            string variant2 = TestContext.Properties["MenuVariantACE4"].ToString();
            ///string variant2 = "BASAL COLEGIO - ALMUERZO";

            Random rnd = new Random();

            string menuName = menuNameToday + "-" + rnd.Next().ToString();
            string menuName1 = menuNameToday + "-" + rnd.Next().ToString();

            string guestType1 = variant1.Substring(0, variant1.IndexOf("-")).Trim();
            string guestType2 = variant2.Substring(0, variant2.IndexOf("-")).Trim();

            //Arrange        
            HomePage homePage = LogInAsAdmin();

            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant1);
            menusPage = menuDayViewPage.BackToList();

            menusCreateModalPage = menusPage.MenuCreatePage();
            menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName1, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant2);
            menusPage = menuDayViewPage.BackToList();

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            menusPage.Filter(MenusPage.FilterType.TypeGuest, guestType1);
            menusPage.UnselectServicesFromMenu();
            Assert.AreEqual(1, menusPage.CheckTotalNumber(), "Aucun menu n'est associé au type de guest sélectionné.");
            Assert.AreEqual(menusPage.GetFirstMenuName(), menuName, String.Format(MessageErreur.FILTRE_ERRONE, "Type Guest"));

            WebDriver.Navigate().Refresh();

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            menusPage.Filter(MenusPage.FilterType.TypeGuest, guestType2);
            menusPage.UnselectServicesFromMenu();
            Assert.AreEqual(0, menusPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "Type Guest"));

            WebDriver.Navigate().Refresh();

            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName1);
            menusPage.Filter(MenusPage.FilterType.TypeGuest, guestType2);
            menusPage.UnselectServicesFromMenu();
            Assert.AreEqual(1, menusPage.CheckTotalNumber(), "Aucun menu n'est associé au type de guest sélectionné.");
            Assert.AreEqual(menusPage.GetFirstMenuName(), menuName1, String.Format(MessageErreur.FILTRE_ERRONE, "Type guest"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Filter_TypeMeal()
        {
            //Prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string variant1 = TestContext.Properties["MenuVariantACE1"].ToString();
            Random rnd = new Random();

            string menuName = menuNameToday + "-" + rnd.Next().ToString();

            string mealType1 = variant1.Substring(variant1.IndexOf("-") + 1).Trim();

            //Arrange
            HomePage homePage = LogInAsAdmin();   
            var menusPage = homePage.GoToMenus_Menus();
            try
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant1);
                menusPage = menuDayViewPage.BackToList();

                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
                menusPage.Filter(MenusPage.FilterType.TypeMeal, mealType1);
                menusPage.UnselectServicesFromMenu();
                if (!menusPage.isPageSizeEqualsTo100())
                {
                    menusPage.PageSize("8");
                    menusPage.PageSize("100");
                }
                var verifyMeal = menusPage.VerifyMeal(mealType1);
                Assert.IsTrue(verifyMeal, String.Format(MessageErreur.FILTRE_ERRONE, "Type of Meal"));
            }

            finally 
            {
                menusPage.MassiveDeleteMenus(menuName, site, variant1);
            }

           
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_ModifyGeneralInformations()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            string newCommercialName = "CommercialName";
            double newBudgetMax = 10;
            double newSalesPrice = 10;
            double newNumberPAX = 5;
            string newCalulationMethod = "STD";

            //Arrange
            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            var menusGeneralInformations = menuDayViewPage.ClickOnGeneralInformation();

            // On remplace les valeurs
            var commercialName = menusGeneralInformations.GetCommercialName();
            menusGeneralInformations.SetCommercialName(newCommercialName);

            var budgetmax = menusGeneralInformations.GetBudgetMax(decimalSeparatorValue);
            menusGeneralInformations.SetBudgetMax(newBudgetMax.ToString());

            var numberPAX = menusGeneralInformations.GetNumberPAX();
            menusGeneralInformations.SetNumberPAX(newNumberPAX.ToString());

            var salesPrice = menusGeneralInformations.GetSalesPrice(decimalSeparatorValue);
            menusGeneralInformations.SetSalesPrice(newSalesPrice.ToString());

            var calculationMethod = menusGeneralInformations.GetCalculationMethod();
            menusGeneralInformations.SetCalculationMethod(newCalulationMethod);

            menusPage = menusGeneralInformations.BackToList();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);

            menuDayViewPage = menusPage.SelectFirstMenu();
            menusGeneralInformations = menuDayViewPage.ClickOnGeneralInformation();

            // Assert
            Assert.AreEqual(newCommercialName, menusGeneralInformations.GetCommercialName(), "Le commercial name n'a pas été pris en compte.");
            Assert.AreNotEqual(newCommercialName, commercialName, "Le commercial name n'a pas été modifié.");

            Assert.AreEqual(newBudgetMax, menusGeneralInformations.GetBudgetMax(decimalSeparatorValue), "Le budget max n'a pas été pris en compte.");
            Assert.AreNotEqual(newBudgetMax, budgetmax, "Le budget max n'a pas été modifié.");

            Assert.AreEqual(newNumberPAX, menusGeneralInformations.GetNumberPAX(), "Le number PAX n'a pas été pris en compte.");
            Assert.AreNotEqual(newNumberPAX, numberPAX, "Le number PAX n'a pas été modifié.");

            Assert.AreEqual(newSalesPrice, menusGeneralInformations.GetSalesPrice(decimalSeparatorValue), "Le sales price n'a pas été pris en compte.");
            Assert.AreNotEqual(newSalesPrice, salesPrice, "Le sales price n'a pas été modifié.");

            Assert.AreEqual(newCalulationMethod, menusGeneralInformations.GetCalculationMethod(), "La calculation method n'a pas été prise en compte.");
            Assert.AreNotEqual(newCalulationMethod, calculationMethod, "La calculation method n'a pas été modifiée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_ChangeDay_Bandeau()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            MenusPage menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            MenusCreateModalPage menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now.AddDays(-1), DateUtils.Now.AddMonths(1), site, variant, serviceName);
            DateTime jourInitial = menuDayViewPage.GetDayDisplayed();

            Assert.AreEqual(jourInitial, DateUtils.Now.Date, "La date du jour n'est pas affichée pour le menu.");

            menuDayViewPage.ClickOnDayAfter();
            DateTime jourFinal = menuDayViewPage.GetDayDisplayed();

            //Assert
            Assert.AreNotEqual(jourInitial, jourFinal, "Le jour visualisé n'a pas été modifié en cliquant sur la flèche 'jour suivant'.");
            Assert.AreEqual(jourFinal, DateUtils.Now.AddDays(1).Date, "La date de demain n'est pas affichée pour le menu.");

            menuDayViewPage.ClickOnDayBefore();
            jourFinal = menuDayViewPage.GetDayDisplayed();

            //Assert
            Assert.AreEqual(jourInitial, jourFinal, "Le jour visualisé n'a pas été modifié en cliquant sur la flèche 'jour précédent'.");
            Assert.AreEqual(jourFinal, DateUtils.Now.Date, "La date du jour n'est pas affichée de nouveau pour le menu.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_ChangeDay_ShowMenuDay()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            MenusPage menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            MenusCreateModalPage menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now.AddDays(-1), DateUtils.Now.AddMonths(1), site, variant, serviceName);
            DateTime jourInitial = menuDayViewPage.GetDayDisplayed();

            Assert.AreEqual(jourInitial, DateUtils.Now.Date, "La date du jour n'est pas affichée pour le menu.");

            menuDayViewPage.SetShowMenuDay(DateUtils.Now.AddDays(1));
            DateTime jourFinal = menuDayViewPage.GetDayDisplayed();

            //Assert
            Assert.AreNotEqual(jourInitial, jourFinal, "Le jour visualisé n'a pas été modifié en changeant la date dans 'Show menu day'.");
            Assert.AreEqual(jourFinal, DateUtils.Now.AddDays(1).Date, "La date de demain n'est pas affichée pour le menu.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_ChangeDay_PlanningOfMonth()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now.AddDays(-1), DateUtils.Now.AddMonths(1), site, variant, serviceName);
            var jourInitial = menuDayViewPage.GetDayDisplayed();

            Assert.AreEqual(jourInitial, DateUtils.Now.Date, "La date du jour n'est pas affichée pour le menu.");

            //Changer de jour via la zone "menu planning of month"
            bool isDayFound = menuDayViewPage.SetMenuPlanningOfMonth(DateUtils.Now.AddDays(1));
            if (!isDayFound)
            {
                //Changer de jour via la zone "show menu day"
                menuDayViewPage.SetShowMenuDay(DateUtils.Now.AddDays(2));
                menuDayViewPage.SetMenuPlanningOfMonth(DateUtils.Now.AddDays(1));
            }

            var jourFinal = menuDayViewPage.GetDayDisplayed();

            //Assert
            Assert.AreNotEqual(jourInitial, jourFinal, "Le jour visualisé n'a pas été modifié en changeant la date dans 'Planning of month'.");
            Assert.AreEqual(jourFinal, DateUtils.Now.AddDays(1).Date, "La date de demain n'est pas affichée pour le menu.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_AddRecipe_DayView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            HomePage homePage = LogInAsAdmin();         

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");

                recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
            }

            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);

            menuDayViewPage.AddRecipe(recipeName);

            //Assert
            Assert.IsTrue(menuDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été ajoutée au menu.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_EditMenu_DayView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";
            string recipeMethod = "Std";
            string recipeCoef = "60";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange            
            var homePage = LogInAsAdmin();
            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            try
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);

                menuDayViewPage.AddRecipe(recipeName);
                Assert.IsTrue(menuDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été ajoutée au menu.");

                var initMethod = menuDayViewPage.GetRecipeMethod();
                var initCoef = menuDayViewPage.GetRecipeCoef();

                menuDayViewPage.ClickOnFirstRecipe();
                menuDayViewPage.SetRecipeMethod(recipeMethod);
                menuDayViewPage.SetRecipeCoef(recipeCoef);

                var menusWeekViewPage = menuDayViewPage.SwitchToWeekView();
                Assert.IsTrue(menusWeekViewPage.IsWeekViewVisible(), "La vue 'Week vue' n'est pas accessible sur l'écran.");

                menuDayViewPage = menusWeekViewPage.SwitchToDayView();

                var finalMethod = menuDayViewPage.GetRecipeMethod();
                var finalCoef = menuDayViewPage.GetRecipeCoef();

                Assert.AreNotEqual(initCoef, finalCoef, "La valeur du coef de la recette du menu n'a pas été modifiée.");
                Assert.AreNotEqual(initMethod, finalMethod, "La valeur de la méthode de calcul de la recette du menu n'a pas été modifiée.");
            }
            finally
            {
                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_UpdatePAXDayValue()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            MenusDayViewPage menuDayViewPage;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Acttry

            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            try
            {

                AddServiceForMenu(homePage, serviceName, guestType);
                menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menuDayViewPage.AddRecipe(recipeName);
                Assert.IsTrue(menuDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été ajoutée au menu.");

                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
                menuDayViewPage = menusPage.SelectFirstMenu();

                menuDayViewPage.AddRecipe(recipeName);
                var initPAXDay = menuDayViewPage.GetTheoricalPAXDay();
                menuDayViewPage.SetTheoricalPAXDay((Double.Parse(initPAXDay) + 1).ToString());

                WebDriver.Navigate().Refresh();

                var finalPAXDay = menuDayViewPage.GetTheoricalPAXDay();
                Assert.AreNotEqual(initPAXDay, finalPAXDay, "La valeur du Theorical PAX/Day du menu n'a pas été modifiée.");
            }

            finally
            {
                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_DeleteRecipe_DayView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string recipeName = "RecipeForTestMenu";
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);

            menuDayViewPage.AddRecipe(recipeName);
            bool isRecipeDisplayed = menuDayViewPage.IsRecipeDisplayed();
            Assert.IsTrue(isRecipeDisplayed, "La recette n'a pas été ajoutée au menu.");

            WebDriver.Navigate().Refresh();

            menuDayViewPage.DeleteFirstRecipe();
            bool isRecipeDisplayedAfterDelete = menuDayViewPage.IsRecipeDisplayed();
            Assert.IsFalse(isRecipeDisplayedAfterDelete, "La recette n'a pas été supprimée du menu.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_DuplicateDay_DayView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            DateTime date = DateUtils.Now;
            DateTime otherDate = DateUtils.Now.AddDays(1);

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, date, DateUtils.Now.AddMonths(1), site, variant, serviceName);

            menuDayViewPage.AddRecipe(recipeName);
            bool isRecipeDisplayed = menuDayViewPage.IsRecipeDisplayed();
            Assert.IsTrue(isRecipeDisplayed, "La recette n'a pas été ajoutée au menu.");

            menuDayViewPage.SetShowMenuDay(otherDate);
            Assert.IsFalse(menuDayViewPage.IsRecipeDisplayed(), "Une recette est déjà présente pour le jour sélectionné.");

            menuDayViewPage.SetShowMenuDay(DateUtils.Now);
            menuDayViewPage.DuplicateMenuDay(otherDate);

            WebDriver.Navigate().Refresh();
            menuDayViewPage.SetShowMenuDay(otherDate);
            Assert.IsTrue(menuDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été dupliquée sur un autre jour.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_SwitchToWeekView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            MenusDayViewPage menuDayViewPage;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            if (menusPage.CheckTotalNumber() == 0)
            {
                AddServiceForMenu(homePage, serviceName, guestType);
                menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menuDayViewPage.AddRecipe(recipeName);
                Assert.IsTrue(menuDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été ajoutée au menu.");

                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            }

            menuDayViewPage = menusPage.SelectFirstMenu();

            var menusWeekViewPage = menuDayViewPage.SwitchToWeekView();
            Assert.IsTrue(menusWeekViewPage.IsWeekViewVisible(), "La vue 'Week vue' n'est pas accessible sur l'écran.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_ChangeWeek_WeekView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";
            string recipeNameBis = "Recipe2ForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            DateTime date = DateUtils.Now;
            DateTime otherDate = date.AddDays(7);

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            MenusPage menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            MenusCreateModalPage menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            MenusWeekViewPage menuWeekViewPage = menuDayViewPage.SwitchToWeekView();

            menuWeekViewPage.AddRecipe(recipeName, date);
            menuWeekViewPage.AddRecipe(recipeNameBis, otherDate);

            bool isRecipeDisplayed = menuWeekViewPage.IsRecipeDisplayed(date, recipeName);
            Assert.IsTrue(isRecipeDisplayed, $"La recette { recipeName} n'a pas été ajoutée.");
            isRecipeDisplayed = menuWeekViewPage.IsRecipeDisplayed(otherDate, recipeNameBis);
            Assert.IsFalse(isRecipeDisplayed, $"La recette { recipeNameBis} est visible pour la semaine en cours.");

            menuWeekViewPage.ClickOnNextWeek();
            isRecipeDisplayed = menuWeekViewPage.IsRecipeDisplayed(otherDate, recipeNameBis);
            Assert.IsTrue(isRecipeDisplayed, $"La recette { recipeNameBis} n'est pas visible pour la semaine prochaine.");
            isRecipeDisplayed = menuWeekViewPage.IsRecipeDisplayed(date, recipeName);
            Assert.IsFalse(isRecipeDisplayed, $"La recette { recipeName} est visible pour la semaine prochaine.");

            menuWeekViewPage.ClickOnPreviousWeek();
            isRecipeDisplayed = menuWeekViewPage.IsRecipeDisplayed(date, recipeName);
            Assert.IsTrue(isRecipeDisplayed, $"La recette { recipeName} n'est pas visible pour la semaine en cours.");
            isRecipeDisplayed = menuWeekViewPage.IsRecipeDisplayed(otherDate, recipeNameBis);
            Assert.IsFalse(isRecipeDisplayed, $"La recette { recipeNameBis} est visible pour la semaine en cours.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_DeleteRecipe_WeekView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            DateTime date = DateUtils.Now;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            MenusPage menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            MenusCreateModalPage menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            MenusWeekViewPage menuWeekViewPage = menuDayViewPage.SwitchToWeekView();

            menuWeekViewPage.AddRecipe(recipeName, date);
            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(date, recipeName), "La recette n'a pas été ajoutée.");

            menuWeekViewPage.ClickOnRecipe(date, recipeName);
            menuWeekViewPage.DeleteRecipe();

            WebDriver.Navigate().Refresh();
            Assert.IsFalse(menuWeekViewPage.IsRecipeDisplayed(date, recipeName), "La recette n'a pas été supprimée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_DuplicateRecipe_WeekView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            DateTime date = DateUtils.Now;
            DateTime otherDate = date.AddDays(1);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menuDayViewPage.WaitPageLoading();
            var menuWeekViewPage = menuDayViewPage.SwitchToWeekView();
            menuDayViewPage.WaitPageLoading();

            double numJour = menuWeekViewPage.GetDayForAddedRecipe(date);
            bool changeVueOtherDay = false;

            if (numJour == 6)
            {
                changeVueOtherDay = true;
            }

            menuWeekViewPage.AddRecipe(recipeName, date);
            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(date, recipeName), "La recette n'a pas été ajoutée.");
            Assert.IsFalse(menuWeekViewPage.IsRecipeDisplayed(otherDate, recipeName, changeVueOtherDay), "La recette existe déjà pour le jour sur lequel la duplication doit être effectuée.");
            menuDayViewPage.WaitPageLoading();

            menuWeekViewPage.ClickOnRecipe(date, recipeName);
            menuWeekViewPage.DuplicateRecipe(otherDate);

            WebDriver.Navigate().Refresh();
            var resultat = menuWeekViewPage.IsRecipeDisplayed(date, recipeName);
            var recetteDuplique = menuWeekViewPage.IsRecipeDisplayed(otherDate, recipeName, changeVueOtherDay);
            //assert 
            Assert.IsTrue(resultat, "La recette a été supprimée du jour initial.");
            Assert.IsTrue(recetteDuplique, "La recette n'a pas été dupliquée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_DuplicateMultiRecipes_WeekView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";
            string recipeNameBis = "Recipe2ForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            DateTime date = DateUtils.Now;
            DateTime otherDate = date.AddDays(1);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            var menuWeekViewPage = menuDayViewPage.SwitchToWeekView();

            double numJour = menuWeekViewPage.GetDayForAddedRecipe(date);
            bool changeVueOtherDay = false;

            if (numJour == 6)
            {
                changeVueOtherDay = true;
            }

            menuWeekViewPage.AddRecipe(recipeName, date);
            menuWeekViewPage.AddRecipe(recipeNameBis, date);

            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(date, recipeName), "La recette " + recipeName + " n'a pas été ajoutée.");
            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(date, recipeNameBis), "La recette " + recipeNameBis + " n'a pas été ajoutée.");

            menuWeekViewPage.ClickOnRecipe(date, recipeName);
            menuWeekViewPage.SelectMultiRecipes();

            // Sélection des recettes à dupliquer
            menuWeekViewPage.ClickOnRecipe(date, recipeNameBis);
            menuWeekViewPage.DuplicateMultiRecipe(otherDate);

            WebDriver.Navigate().Refresh();
            menuWeekViewPage.WaitPageLoading();
            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(date, recipeName), "La recette " + recipeName + " a été supprimée.");
            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(date, recipeNameBis), "La recette " + recipeNameBis + " a été supprimée.");
            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(otherDate, recipeName, changeVueOtherDay), "La recette " + recipeName + " n'a pas été dupliquée.");
            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(otherDate, recipeNameBis, changeVueOtherDay), "La recette " + recipeName + " n'a pas été dupliquée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_DeleteMultiRecipes_WeekView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";
            string recipeNameBis = "Recipe2ForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            DateTime date = DateUtils.Now;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            var menuWeekViewPage = menuDayViewPage.SwitchToWeekView();

            menuWeekViewPage.AddRecipe(recipeName, date);
            menuWeekViewPage.AddRecipe(recipeNameBis, date);

            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(date, recipeName), "La recette " + recipeName + " n'a pas été ajoutée.");
            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(date, recipeNameBis), "La recette " + recipeNameBis + " n'a pas été ajoutée.");

            menuWeekViewPage.ClickOnRecipe(date, recipeName);
            menuWeekViewPage.SelectMultiRecipes();

            // Sélection des recettes à dupliquer
            menuWeekViewPage.ClickOnRecipe(date, recipeNameBis);
            menuWeekViewPage.DeleteMultiRecipes();

            WebDriver.Navigate().Refresh();
            Assert.IsFalse(menuWeekViewPage.IsRecipeDisplayed(date, recipeName), "La recette " + recipeName + " n'a pas été supprimée.");
            Assert.IsFalse(menuWeekViewPage.IsRecipeDisplayed(date, recipeNameBis), "La recette " + recipeNameBis + " n'a pas été supprimée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_AddRecipe_WeekView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            DateTime date = DateUtils.Now;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            var menuWeekViewPage = menuDayViewPage.SwitchToWeekView();

            menuWeekViewPage.AddRecipe(recipeName, date);

            //Assert
            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(date, recipeName), "La recette n'a pas été ajoutée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_ModifyRecipe_WeekView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            DateTime date = DateUtils.Now;
            string newMethod = "Standard";
            string newCoef = "82";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            var menuWeekViewPage = menuDayViewPage.SwitchToWeekView();

            menuWeekViewPage.AddRecipe(recipeName, date);
            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(date, recipeName), "La recette n'a pas été ajoutée.");

            menuWeekViewPage.ClickOnRecipe(date, recipeName);
            string initCoef = menuWeekViewPage.GetRecipeCoefficient();
            string initMethodValue = menuWeekViewPage.GetRecipeMethodValue();

            // Modification des valeurs
            menuWeekViewPage.SetRecipeMethod(newMethod);
            menuWeekViewPage.SetRecipeCoefficient(newCoef);
            menuDayViewPage.WaitLoading();
            menuWeekViewPage.CloseModale();

            WebDriver.Navigate().Refresh();
            menuWeekViewPage.ClickOnRecipe(date, recipeName);

            var Coef = menuWeekViewPage.GetRecipeCoefficient();
            var MethodValue = menuWeekViewPage.GetRecipeMethodValue();
            // Assert
            Assert.AreNotEqual(initCoef, Coef, "Le coefficient de la recette n'a pas été modifié.");
            Assert.AreNotEqual(initMethodValue, MethodValue, "La méthode de la recette n'a pas été modifiée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_EditRecipe_WeekView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            DateTime date = DateUtils.Now;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            var menuWeekViewPage = menuDayViewPage.SwitchToWeekView();

            menuWeekViewPage.AddRecipe(recipeName, date);
            Assert.IsTrue(menuWeekViewPage.IsRecipeDisplayed(date, recipeName), "La recette n'a pas été ajoutée.");

            menuWeekViewPage.ClickOnRecipe(date, recipeName);
            var recipeInformationPage = menuWeekViewPage.EditRecipe();

            try
            {
                recipeInformationPage.ClickOnEditInformation();
                Assert.AreEqual(recipeName, recipeInformationPage.GetRecipeName(), "L'onglet de la recette ne s'est pas ouvert.");
            }
            finally
            {
                recipeInformationPage.Close();
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_SwitchToDayView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string recipeName = "RecipeForTestMenu";
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            MenusDayViewPage menuDayViewPage;

            //Arrange
            var homePage = LogInAsAdmin();
            
            //Act
            AddServiceForMenu(homePage, serviceName, guestType);
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menuDayViewPage.AddRecipe(recipeName);
            bool isRecipeDisplayed = menuDayViewPage.IsRecipeDisplayed();
            Assert.IsTrue(isRecipeDisplayed, "La recette n'a pas été ajoutée au menu.");

            var menusWeekViewPage = menuDayViewPage.SwitchToWeekView();
            bool isWeekViewVisible = menusWeekViewPage.IsWeekViewVisible();
            Assert.IsTrue(isWeekViewVisible, "La vue 'Week view' n'est pas accessible sur l'écran.");

            menuDayViewPage = menusWeekViewPage.SwitchToDayView();
            bool isRecipeDisplayedSwitchToDayView = menuDayViewPage.IsRecipeDisplayed();
            Assert.IsTrue(isRecipeDisplayedSwitchToDayView, "La vue 'Day view' n'est pas affichée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_ExportToAnotherMenu()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            Random rnd = new Random();

            string menuName = menuNameToday + "-" + rnd.Next().ToString();
            string menuName2 = menuNameToday + "-" + rnd.Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            MenusDayViewPage menuDayViewPage;

            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");

                recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
            }
            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);

            Assert.IsFalse(menuDayViewPage.IsRecipeDisplayed(), "Le menu possède déjà des recettes.");
            menusPage = menuDayViewPage.BackToList();

            menusCreateModalPage = menusPage.MenuCreatePage();
            menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName2, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menuDayViewPage.AddRecipe(recipeName);
            Assert.IsTrue(menuDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été ajoutée au menu.");

            var menusGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
            menusGeneralInformationPage.ExportToOtherMenu(site, menuName);
            Assert.IsTrue(menusGeneralInformationPage.IsExportOK(), "L'export des recettes vers un autre menu s'est terminée en erreur.");
            menusPage = menuDayViewPage.BackToList();

            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            menuDayViewPage = menusPage.SelectFirstMenu();

            menuDayViewPage.SetShowMenuDay(DateUtils.Now);
            Assert.IsTrue(menuDayViewPage.IsRecipeDisplayed(), "La recette a été exportée vers l'autre menu.");
            Assert.AreEqual(menuDayViewPage.GetRecipeName(), recipeName, "La recette " + recipeName + " a bien été copiée dans l'autre menu.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Print_Menu_NewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeName = "RecipeForTestMenu";

            bool versionPrint = true;

            MenusDayViewPage menuDayViewPage;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            var menusPage = homePage.GoToMenus_Menus();

            menusPage.ClearDownloads();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);

            menuDayViewPage.AddRecipe(recipeName);
            Assert.IsTrue(menuDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été ajoutée au menu.");

            var menuGeneralInformation = menuDayViewPage.ClickOnGeneralInformation();
            var reportPage = menuGeneralInformation.Print(versionPrint);
            var isGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isGenerated, "Le menu n'a pas été imprimé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Export_Menu_NewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeName = "RecipeForTestMenu";
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            bool versionPrint = true;

            MenusDayViewPage menuDayViewPage;

            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");

                recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
            }

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            var menusPage = homePage.GoToMenus_Menus();

            menusPage.ClearDownloads();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);

            menuDayViewPage.AddRecipe(recipeName);
            Assert.IsTrue(menuDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été ajoutée au menu.");

            var menuWeekViewPage = menuDayViewPage.SwitchToWeekView();

            DeleteAllFileDownload();

            menuWeekViewPage.Export(versionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = menuWeekViewPage.GetExportExcelFile(taskFiles, menuName.Replace("/", "_"));
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber(MENUS_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Menu", MENUS_EXCEL_SHEET_NAME, filePath, menuName);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, "Le fichier imprimé contient des menus différents que celui en cours.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_DeleteEntireMenu()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            //Créer un menu avec un variant et un service
            var menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);

            //1. Dans General Information, Clear Services --> Vérifier que le service est bien enlevé
            MenusGeneralInformationPage generalInfo = menuDayViewPage.ClickOnGeneralInformation();
            menuDayViewPage = generalInfo.ClearService();
            //2.Cliquer sur Activated --> Vérifier que c'est bien désactiver
            generalInfo = menuDayViewPage.ClickOnGeneralInformation();
            generalInfo.Activate();
            //3.Cliquer sur Delete entire menu
            MenusPage menuPage = generalInfo.DeleteEntireMenu();
            //Revenir sur l'index et vérifier qu'il a bien été supprimé
            menuPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            Assert.AreEqual(0, menuPage.CheckTotalNumber(), "menu non effacé");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_MassiveDelete()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string recipeName = "RecipeForTestMenu";
            string menuName = menuNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage= LogInAsAdmin();
         
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);

            menuDayViewPage.AddRecipe(recipeName);
            Assert.IsTrue(menuDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été ajoutée au menu.");

            menusPage = menuDayViewPage.BackToList();

            menusPage.MassiveDeleteMenus(menuName, site, variant);

            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);

            //Assert
            Assert.AreEqual(0, menusPage.CheckTotalNumber(), "La massive delete ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_DetailExportToOtherMenu()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeName = "RecipeForTestMenu";
            string menuName2 = "Menu-2" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            //Arrange

            HomePage homePage = LogInAsAdmin();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            var menusPage = homePage.GoToMenus_Menus();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menusDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menusPage = menusDayViewPage.BackToList();
            //create un deuxiéme menu avec recipe
            menusCreateModalPage = menusPage.MenuCreatePage();
            menusDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName2, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);

            menusDayViewPage.AddRecipe(recipeName);
            Assert.IsTrue(menusDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été ajoutée au menu.");
            menusDayViewPage.ExportToOtherMenu(site, menuName);
            menusPage = menusDayViewPage.BackToList();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            menusDayViewPage = menusPage.SelectFirstMenu();
            int nbreRecipes = menusDayViewPage.CheckTotalRecipe();
            var recipe = menusDayViewPage.GetRecipeName();
            var isRecipeExist = menusDayViewPage.VerifyrecipeExist(recipe);
            //Assert
            Assert.AreEqual(1, nbreRecipes, MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE);
            Assert.IsTrue(isRecipeExist, MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_DetailPlanView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeName = "RecipeForTestMenu";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            MenusPage menusPage = homePage.GoToMenus_Menus();

            MenusCreateModalPage menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menusDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menusDayViewPage.AddRecipe(recipeName);
            bool isRecipeDisplayed = menusDayViewPage.IsRecipeDisplayed();
            Assert.IsTrue(isRecipeDisplayed, "La recette n'a pas été ajoutée au menu.");

            menusDayViewPage.ClickPlanView();
            bool isPlanViewChanged = menusDayViewPage.VerifyPlanViewChanged();
            Assert.IsTrue(isPlanViewChanged, "Les lignes des menu ne sont pas changés d'affichage");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_PrintDetail_MenuDayView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeName = "RecipeForTestMenu";
            bool versionPrint = true;
            string DocFileNamePdfBegin = "Menu Report";
            string DocFileNamePdfBegin2 = "Menu Per Week";
            string DocFileNameZipBegin = "All_files_";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            homePage.ClearDownloads();
            AddServiceForMenu(homePage, serviceName, guestType);

            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menusDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menusDayViewPage.AddRecipe(recipeName);
            Assert.IsTrue(menusDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été ajoutée au menu.");

            var printReportPage = menusDayViewPage.Print(versionPrint);
            var isGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();
            //Assert
            Assert.IsTrue(isGenerated, "Le menu n'a pas été imprimé.");
            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin2, DocFileNameZipBegin);

            // cliquer sur All
            string trouve = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(1, words.Where(w => w.Text.Contains(site)).Count(), site + " non présent dans le Pdf");

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_PrintDetail_MenuWeekView()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeName = "RecipeForTestMenu";
            bool versionPrint = true;
            string DocFileNamePdfBegin = "Menu Per Week";
            string DocFileNamePdfBegin2 = "Menu Report";

            string DocFileNameZipBegin = "All_files_";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            MenusPage menusPage = homePage.GoToMenus_Menus();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            MenusDayViewPage menusDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menusDayViewPage.AddRecipe(recipeName);
            Assert.IsTrue(menusDayViewPage.IsRecipeDisplayed(), "La recette n'a pas été ajoutée au menu.");

            MenusWeekViewPage menusWeekViewPage = menusDayViewPage.SwitchToWeekView();
            Assert.IsTrue(menusWeekViewPage.IsWeekViewVisible(), "La vue 'Week vue' n'est pas accessible sur l'écran.");

            PrintReportPage printReportPage = menusWeekViewPage.Print(versionPrint);
            bool isGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();
            Assert.IsTrue(isGenerated, "Le menu n'a pas été imprimé.");

            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin2, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            //Assert
            Assert.AreEqual(1, words.Where(w => w.Text.Contains(recipeName)).Count(), "La recette" + recipeName + " non présente dans le Pdf");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Envoie_des_donnees_Display_EAT()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            string recipeName = "RecipeForTestMenu";
            string recipeNameBis = "Recipe2ForTestMenu";

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            var homePage = LogInAsAdmin();


            AddServiceForMenu(homePage, serviceName, guestType);

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            try
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menuDayViewPage.AddRecipe(recipeName);
                menuDayViewPage.AddRecipe(recipeNameBis);

                var menuDisplayPage = menuDayViewPage.OpenDisplayMenu();

                //Assert
                var checkRecipeName = menuDisplayPage.CheckHasRecipeName();
                var checkRecipePointPrice = menuDisplayPage.CheckHasRecipePointsPrice();
                var checkRecipeSalePrice = menuDisplayPage.CheckHasRecipeSalesPrice();
                var CheckHasRecipeEnergy = menuDisplayPage.CheckHasRecipeEnergy();
                Assert.IsTrue(checkRecipeName, "Aucune recette n'est affichée");
                Assert.IsTrue(checkRecipePointPrice, "Il n'y a pas de point price");
                Assert.IsTrue(checkRecipeSalePrice, "Il n'y a pas de sales price");
                Assert.IsTrue(CheckHasRecipeEnergy, "Il n'y a pas d'énergie");
            }
            finally
            {
                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant); 
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_sitesearch()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var menusPage = homePage.GoToMenus_Menus();

            string site = TestContext.Properties["Site"].ToString();

            menusPage.MassiveDeleteSiteSearch(site);
            Assert.IsTrue(menusPage.VerifySiteInMassiveDeleteSearch(site), "Massive delete search non fonctionnel");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_namesearch()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var menusPage = homePage.GoToMenus_Menus();

            var menuName = menusPage.GetFirstMenuName();
            menusPage.MassiveDeleteNameSearch(menuName);
            Assert.IsTrue(menusPage.VerifyMenuNameInMassiveDeleteSearch(menuName), "Massive delete search non fonctionnel");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_index_CreateValidator()
        {
            HomePage homePage = LogInAsAdmin();     
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            menusCreateModalPage.ClickCreate();
            Assert.IsTrue(menusCreateModalPage.IsMessagesValidatorsDisplayed(), "Les messages validators n'apparaissent pas");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_index_ResetFilter()
        {            
            var homePage = LogInAsAdmin();
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            var defaultName = menusPage.GetNameSearched();
            var defaultActiveOrInactive = menusPage.GetIsActiveOrInactiveChecked();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, "MenuPUBO");
            menusPage.Filter(MenusPage.FilterType.ShowInactive, true);
            var ResultWithFilters = menusPage.CheckTotalNumber();
            menusPage.ResetFilter();
            var ResultWithoutFilters = menusPage.CheckTotalNumber();
            Assert.AreNotEqual(ResultWithFilters, ResultWithoutFilters, "Les filters ne sont pas réinitialisé");
            Assert.IsTrue(defaultName == menusPage.GetNameSearched(), "Le Search filter n'est pas réinitialisé");
            Assert.IsTrue(defaultActiveOrInactive == menusPage.GetIsActiveOrInactiveChecked(), "Le Filter par check n'est pas réinitialisé");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_filter_RecipeType()
        {
            string recipeTypeSelected = TestContext.Properties["RecipeType"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            var menusDayViewPage = menusPage.SelectFirstMenu();
            menusDayViewPage.Filter(MenusDayViewPage.FilterType.SearchByRecipe, "Rec");
            // only recipeType = "PLATO CALIENTE" have data (RECIPE)
            menusDayViewPage.Filter(MenusDayViewPage.FilterType.RecipeTypes, recipeTypeSelected);
            var listRecipesResultFiltered = menusDayViewPage.GetAllRecipeNameResultFiltredByRecipe();
            homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var recipesPage = homePage.GoToMenus_Recipes();

            Assert.IsTrue(listRecipesResultFiltered.Count() > 0, "No Recipe Found");
            foreach (var recipe in listRecipesResultFiltered)
            {
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipe);
                var recipeGeneralInformationPage = recipesPage.SelectFirstRecipe();
                recipeGeneralInformationPage.ClickOnEditInformation();
                var recipeTypeResultFiltred = recipeGeneralInformationPage.GetRecipeType();
                Assert.AreEqual(recipeTypeResultFiltred, recipeTypeSelected, $"le recette {recipe} qui ressort n'a pas bien le recipe type {recipeTypeSelected} sélectionné");
                recipeGeneralInformationPage.BackToList();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_filter_DietaryTtpes()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariantACE1"].ToString();
            string ingredient = "iemForDietaryTtpes";
            string site = TestContext.Properties["SiteACE"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            bool isVariantCreated = false;
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();



            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            try
            {

                ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, ingredient);
                if (itemPage.CheckTotalNumber() == 0)
                {
                    var itemCreateModalPage = itemPage.ItemCreatePage();
                    var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient, group, workshop, taxType, prodUnit);

                    var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                    itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                    itemPage = itemGeneralInformationPage.BackToList();
                }
                // 2 - Création de la recette pour la variante
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString(), true, "Arroz");
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);

                Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

                var menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant);
                menuDayViewPage.SelectDietaryType("Arroz");
                menuDayViewPage.AddRecipe(recipeName);
                var recipeExist = menuDayViewPage.VerifyrecipeExist(recipeName);
                Assert.IsTrue(recipeExist, "La recette n'a pas été ajouté à la menu.");

            }
            finally
            {
                var menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);

                RecipesPage recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_index_CreateNoVariant()
        {

            string site = TestContext.Properties["Site"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
   
            HomePage homePage = LogInAsAdmin();

            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            menusCreateModalPage.FillField_CreateNewMenuNoVariant(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site);
            var variantRequiredMessage = menusCreateModalPage.GetAlertMessage();
            Assert.AreEqual(variantRequiredMessage, "You must select one variant minimum");

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_GeneralInfoAddService()
        {
            //Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceCategory = TestContext.Properties["ServiceCategory"].ToString();
            List<string> servicesToAdd = new List<string> { "Service FoodCost B", "Service FoodCost A" };

            //Arrange
            var homePage = LogInAsAdmin();

            //act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.Site, siteACE);
            var menusDayViewPage = menusPage.SelectFirstMenu();
            MenusGeneralInformationPage menusGeneralInformation = menusDayViewPage.ClickOnGeneralInformation();
            try
            {
                menusGeneralInformation.AddServices(serviceCategory, servicesToAdd);
                var servicesAdded = menusGeneralInformation.GetServicesAddedToMenu();
                foreach (var service in servicesToAdd)
                {
                    Assert.IsTrue(servicesAdded.Contains(service.Trim()), $"Le service {service} n'est pas ajouté");
                }
            }
            finally
            {
                menusGeneralInformation.ClearService();
            }

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_GeneralInfoAddServicePopup()
        {
            //Prepare
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string serviceCategory = TestContext.Properties["CategoryService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.Site, siteACE);
            var menusDayViewPage = menusPage.SelectFirstMenu();
            MenusGeneralInformationPage menusGeneralInformation = menusDayViewPage.ClickOnGeneralInformation();
            var allServicesOptions = menusGeneralInformation.GetAllServicesOptionsByCategory(serviceCategory);
           
            homePage.Navigate();
            var servicesPage = homePage.GoToCustomers_ServicePage();
            foreach (var service in allServicesOptions)
            {
                servicesPage.ResetFilters();
                servicesPage.Filter(ServicePage.FilterType.Sites, siteACE);
                servicesPage.Filter(ServicePage.FilterType.Search, service);
                
                //Assert
                Assert.IsTrue(servicesPage.VerifyCategory(serviceCategory), $"le service {service}  qui apparaît dans la liste d'options ne correspond pas au Category {serviceCategory} sélectionné");
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_GeneralInfoClearService()
        {
            string siteMAD = TestContext.Properties["Site"].ToString();
            string serviceCategory = TestContext.Properties["ServiceCategory"].ToString();
            List<string> services = new List<string> { "Service-11/13/2020", "Service-11/13/2021" };
            var homePage = LogInAsAdmin();
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.Site, siteMAD);
            var menusDayViewPage = menusPage.SelectFirstMenu();
            MenusGeneralInformationPage menusGeneralInformation = menusDayViewPage.ClickOnGeneralInformation();
            var IsMenuHaveServices = menusGeneralInformation.GetServicesAddedToMenu();
            if (IsMenuHaveServices.Count() == 0)
            {
                menusGeneralInformation.AddServices(serviceCategory, services);
            }
            menusGeneralInformation.ClearService();
            menusGeneralInformation = menusDayViewPage.ClickOnGeneralInformation();
            IsMenuHaveServices = menusGeneralInformation.GetServicesAddedToMenu();
            Assert.IsTrue(IsMenuHaveServices.Count() == 0, "Les services ne sont pas supprimés du menu");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_GeneralInfoStart()
        {
            //Prepare
            Random rnd = new Random();
            string menuName = "Menu -" + rnd.Next().ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            bool isActive = true;
            string calculationMethodToAdd = "STD";
           
            //Arrange
            var homePage = LogInAsAdmin();

            //act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now.AddDays(-10), DateUtils.Now.AddDays(+30), siteMAD, variant, serviceName, isActive, calculationMethodToAdd);
                menuDayViewPage.BackToList();
            }
            var menusDayViewPage = menusPage.SelectFirstMenu();
            var generalInfos = menusDayViewPage.ClickOnGeneralInformation();
            //Assert 
            var startDateDisabled = generalInfos.IsStartDateDisabled();
            var endDateDisabled = generalInfos.IsEndDateDisabled();
            Assert.IsTrue(startDateDisabled, "Le start date doit etre non modifiable");
            Assert.IsFalse(endDateDisabled, "Le end date doit etre modifiable");
          
            //Delete Menu 
            menusPage = generalInfos.BackToList();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.Filter(MenuMassiveDeletePage.FilterType.SearchMenu, menuName);
            menuMassiveDeletePage.ClickOnSearch();
            menuMassiveDeletePage.ClickOnFirstLine();
            menuMassiveDeletePage.ClickOnDelete();
            menuMassiveDeletePage.ClickOnConfirmDelete();
            menuMassiveDeletePage.ClickOnConfirmAfterDelete();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);

            //Assert
            int numberokDelete = menusPage.CheckTotalNumber();
            Assert.AreEqual(0, numberokDelete, "Menu supprimé avec succès");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_GeneralInfoEnd()
        {
            //Prepare
            Random rnd = new Random();
            string menuName = "Menu -" + rnd.Next().ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            bool isActive = true;
            string calculationMethodToAdd = "STD";
            
            //Arrange
            var homePage = LogInAsAdmin();

            //act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now.AddDays(-10), DateUtils.Now.AddDays(+30), siteMAD, variant, serviceName, isActive, calculationMethodToAdd);
                menuDayViewPage.BackToList();
            }
            var menusDayViewPage = menusPage.SelectFirstMenu();
            var generalInfos = menusDayViewPage.ClickOnGeneralInformation();

            //Assert 
            var startDateDisabled = generalInfos.IsStartDateDisabled();
            var endDateDisabled = generalInfos.IsEndDateDisabled();
            Assert.IsTrue(startDateDisabled, "Le start date doit etre non modifiable");
            Assert.IsFalse(endDateDisabled, "Le end date doit etre modifiable");
            
            //Delete Menu 
            menusPage = generalInfos.BackToList();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.Filter(MenuMassiveDeletePage.FilterType.SearchMenu, menuName);
            menuMassiveDeletePage.ClickOnSearch();
            menuMassiveDeletePage.ClickOnFirstLine();
            menuMassiveDeletePage.ClickOnDelete();
            menuMassiveDeletePage.ClickOnConfirmDelete();
            menuMassiveDeletePage.ClickOnConfirmAfterDelete();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);

            //Assert
            int numberokDelete = menusPage.CheckTotalNumber(); 
            Assert.AreEqual(0, numberokDelete, "Menu supprimé avec succès");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_GeneralInfoDateWeek()
        {

            //Prepare
            Random rnd = new Random();
            string menuName = "Menu -" + rnd.Next().ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            bool isActive = true;
            string calculationMethodToAdd = "STD";

            //Arrange
            var homePage = LogInAsAdmin();

            //act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);

            if (menusPage.CheckTotalNumber() == 0)
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now.AddDays(-10), DateUtils.Now.AddDays(+30), siteMAD, variant, serviceName, isActive, calculationMethodToAdd);
                menuDayViewPage.BackToList();
            }

            var menusDayViewPage = menusPage.SelectFirstMenu();
            var generalInfos = menusDayViewPage.ClickOnGeneralInformation();
            
            //Assert
            var dateOfWeekDisabled = generalInfos.IsDateOfWeekDisabled();
            Assert.IsTrue(dateOfWeekDisabled, "Le date of week doit etre non modifiable");
          
            //Delete Menu 
            menusPage = generalInfos.BackToList();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.Filter(MenuMassiveDeletePage.FilterType.SearchMenu, menuName);
            menuMassiveDeletePage.ClickOnSearch();
            menuMassiveDeletePage.ClickOnFirstLine();
            menuMassiveDeletePage.ClickOnDelete();
            menuMassiveDeletePage.ClickOnConfirmDelete();
            menuMassiveDeletePage.ClickOnConfirmAfterDelete();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);

            //Assert
            int numberokDelete = menusPage.CheckTotalNumber();
            Assert.AreEqual(0, numberokDelete, "Menu supprimé avec succès");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_GeneralInfoActivated()
        {
            var homePage = LogInAsAdmin();
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.Service, "none");
            var menuName = menusPage.GetFirstMenuName();
            var menusDayViewPage = menusPage.SelectFirstMenu();
            MenusGeneralInformationPage menusGeneralInformation = menusDayViewPage.ClickOnGeneralInformation();
            try
            {
                menusGeneralInformation.WaitLoading();
                menusGeneralInformation.Activate();
                menusPage = menusGeneralInformation.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
                menusPage.Filter(MenusPage.FilterType.ShowInactive, true);
                bool checkTotalNumber = (menusPage.CheckTotalNumber() > 0);
                bool firstMenuName = menusPage.GetFirstMenuName() == menuName;
                Assert.IsTrue(checkTotalNumber && firstMenuName, "Le menu n'est pas désactivé");
            }
            finally
            {
                menusDayViewPage = menusPage.SelectFirstMenu();
                menusGeneralInformation = menusDayViewPage.ClickOnGeneralInformation();
                menusGeneralInformation.Activate();
                menusGeneralInformation.BackToList();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_index_pagination()
        {

            
            var homePage = LogInAsAdmin();

            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            var menuNamePage1 = menusPage.GetFirstMenuName();
            menusPage.PageSize("8");
            var numberOfMenus = menusPage.GetNumberOfMenus();

            Assert.IsTrue(numberOfMenus == 8);

            menusPage.GoToPageTwo();
            var menuNamePage2 = menusPage.GetFirstMenuName();

            Assert.AreNotEqual(menuNamePage1, menuNamePage2);

            menusPage.PageSize("100");
            menuNamePage1 = menusPage.GetFirstMenuName();
            numberOfMenus = menusPage.GetNumberOfMenus();

            Assert.IsTrue(numberOfMenus == 100);

            menusPage.GoToPageTwo();
            menuNamePage2 = menusPage.GetFirstMenuName();
            Assert.AreNotEqual(menuNamePage1, menuNamePage2);
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_index_CreateCalculationMethod()
        {
            Random rnd = new Random();
            string menuName = "Menu -" + rnd.Next().ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            bool isActive = true;
            string calculationMethodToAdd = "STD";
            HomePage homePage = LogInAsAdmin();
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now.AddDays(-10), DateUtils.Now.AddDays(+30), siteMAD, variant, serviceName, isActive, calculationMethodToAdd);
            MenusGeneralInformationPage menusGeneralInformation = menuDayViewPage.ClickOnGeneralInformation();
            var calculationMethodAdded = menusGeneralInformation.GetCalculationMethod();
            Assert.AreEqual(calculationMethodToAdd, calculationMethodAdded, "La méthode de calcul ajoutée n'est pas enregistrée");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_index_CreateSiteVariant()
        {
            string site = TestContext.Properties["SitePriceList"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();

            HomePage homePage = LogInAsAdmin();

            ParametersProduction parametersProduction = homePage.GoToParameters_ProductionPage();
            parametersProduction.GoToTab_Variant();
            parametersProduction.FilterSite(site);
            var variantListOfGuests = parametersProduction.GetVariantListOfGuests();
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            var variantsOfSelectedSite = menusCreateModalPage.GetVariantsOfSelectedSite(site);
            var variantsOfSelectedSiteModified = variantsOfSelectedSite.Select(element => element.Split('-')[0].Trim()).ToList();
            bool allExist = variantsOfSelectedSiteModified.All(element => variantListOfGuests.Contains(element));
            Assert.IsTrue(allExist, "les variants porposés dans la liste ne correspondent pas au variants disponibles sur le site sélectionné");
        }
        // _______________________________________ Utilitaire _________________________________________________

        public void AddServiceForMenu(HomePage homePage, string serviceName, string guestType)
        {
            // Prepare
            string customer = TestContext.Properties["MenuServiceCustomer"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string serviceCode = new Random().Next().ToString();
            string DeliveryName = "DeliveryTestForMenu" ;
            string serviceProduction = GenerateName(4);
            ServicePricePage servicePricePage = new ServicePricePage(WebDriver, TestContext);

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            if (servicePage.CheckTotalNumber() == 0)
            {

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction, serviceCategory, guestType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                servicePricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = servicePricePage.AddNewCustomerPrice();
                servicePricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddMonths(2));

            }
            else
            {
                servicePricePage = servicePage.ClickOnFirstService();
                var isExist = servicePricePage.IsPriceExist();
                var priceName = servicePricePage.GetPriceName();
                if (!isExist || priceName != "MAD / AIR CPU SL")
                {
                    var priceModalPage = servicePricePage.AddNewCustomerPrice();
                    servicePricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddMonths(2));
                    servicePage = servicePricePage.BackToList();
                }
                else
                {
                    servicePricePage.UnfoldAll();
                    ServiceCreatePriceModalPage serviceCreatePriceModalPage = servicePricePage.EditFirstPriceService();
                    servicePricePage = serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now, DateUtils.Now.AddMonths(2));
                }

            }

            ServiceDeliveryPage serviceDeliveryPage = servicePricePage.GoToDeliveryTab();
            var deliveriesExist = serviceDeliveryPage.CheckDeliveriesExist();
            if (!deliveriesExist)
            {
                homePage.Navigate();
                var DeliveriesPage = homePage.GoToCustomers_DeliveryPage();
                DeliveriesPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search , DeliveryName);
                if (DeliveriesPage.CheckTotalNumber() == 0)
                {
                    var DeliveryCreateModalPage = DeliveriesPage.DeliveryCreatePage();
                    DeliveryCreateModalPage.FillFields_CreateDeliveryModalPage(DeliveryName, customer, site, true);
                    var deliveryLoading = DeliveryCreateModalPage.Create();
                    deliveryLoading.AddService(serviceName);
                    deliveryLoading.AddQuantityForService("10", serviceName);
                }
                else
                {
                    DeliveryLoadingPage deliveryLoadingPage = DeliveriesPage.ClickOnFirstDelivery();
                    var isVisible =deliveryLoadingPage.IsServiceVisible();
                    if (!isVisible)
                    {
                        deliveryLoadingPage.AddService(serviceName);
                        deliveryLoadingPage.AddQuantityForService("10", serviceName);
                    }
                }


            }

            servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);


            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceName), "Le service " + serviceName + " n'existe pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_CoeffEdit()
        {
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeNameBis = "Recipe2ForTestMenu";
            //Arrange
            
            var homePage = LogInAsAdmin();
            AddServiceForMenu(homePage, serviceName, guestType);
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            var menusPage = homePage.GoToMenus_Menus();
            try
            {

           
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menuDayViewPage.AddRecipe(recipeNameBis);
            var pax_before_change_Coef = menuDayViewPage.GetPAX(decimalSeparatorValue);
            var prixtotal_before_change_Coef = menuDayViewPage.GetPRIXXTOTAL(decimalSeparatorValue);
            menuDayViewPage.ClickOnFirstRecipe();
            menuDayViewPage.SetRecipeCoef("33");
            menusPage = menuDayViewPage.BackToList();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            menuDayViewPage = menusPage.ClickFirstMenu();
            var pax_after_change_Coef = menuDayViewPage.GetPAX(decimalSeparatorValue); ;
            var prixtotal_after_change_Coef = menuDayViewPage.GetPRIXXTOTAL(decimalSeparatorValue);
            Assert.AreNotEqual(pax_before_change_Coef, pax_after_change_Coef, "le prix dans la colonne PAX n est pas met à jour selon les valeurs renseignées.");
            Assert.AreNotEqual(prixtotal_before_change_Coef, prixtotal_after_change_Coef, "Le prix total n est pas met à jour.");
            }
            finally
            {
                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName,site,variant);
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_allergen()
        {
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeName = recipeNameToday + "-" + rnd.Next().ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string keyword = "ECO";
            //Arrange
            var homePage = LogInAsAdmin();
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");

                recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                var itemKeywordPage = recipeGeneralInfosPage.ClickOnKeywordPage();
                itemKeywordPage.AddKeyword(keyword);

            }
            AddServiceForMenu(homePage, serviceName, guestType);
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
            menuGeneralInformationPage.AddConcept();
            menusPage = menuDayViewPage.BackToList();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            menuDayViewPage = menusPage.ClickFirstMenu();
            menuDayViewPage.AddRecipe(recipeName);
            menuDayViewPage.ClickPlanView();
            menuDayViewPage.ClickOnAllergenEtoile();
            menuDayViewPage.CheckFirstAllergenEtoile();
            menuDayViewPage.ClickPlanView();
            WebDriver.Navigate().Refresh();
            menuDayViewPage.ClickPlanView();
            menuDayViewPage.ClickOnAllergenEtoile();
            var IsFirstAllergenChecked = menuDayViewPage.IsFirstAllergenCheked();
            var nbrAllergenChecked = menuDayViewPage.NbAllergenCheked();
            Assert.IsTrue(IsFirstAllergenChecked, "L'information sur l'allergène n'a pas été enregistrée");
            Assert.AreEqual(nbrAllergenChecked, 1, "L'information sur l'allergène n'a pas été enregistrée");


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_DisplayView()
        {
            //Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeName = recipeNameToday + "-" + rnd.Next().ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string keyword = "ECO";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var ProductionPage = homePage.GoToParameters_ProductionPage();
            ProductionPage.GoToTab_Keyword();
            var isKeywordExist = ProductionPage.CheckKeyword(keyword);
            Assert.IsTrue(isKeywordExist, "Se Keyword n'existe pas");
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");
            recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            AddServiceForMenu(homePage, serviceName, guestType);
            var menusPage = homePage.GoToMenus_Menus();
            try
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                menuGeneralInformationPage.AddConcept();
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
                menuDayViewPage = menusPage.ClickFirstMenu();
                menuDayViewPage.AddRecipe(recipeName);
                menuDayViewPage.ClickPlanView();
                //-------------------------------------------------------------------
                menuDayViewPage.ClickOnAllergenEtoile();
                menuDayViewPage.CheckFirstAllergenEtoile();
                //--------------------------------------------------------------------
                menuDayViewPage.ClickOnKeyword();
                menuDayViewPage.CheckKeyword(keyword);
                //---------------------------------------------------------------------
                menuDayViewPage.ClickOnKiosque();
                menuDayViewPage.CheckFirstKiosque();
                //---------------------------------------------------------------------
                WebDriver.Navigate().Refresh();
                //---------------------------------------------------------------------
                menuDayViewPage.ClickPlanView();
                menuDayViewPage.ClickOnAllergenEtoile();
                var IsallergenChecked = menuDayViewPage.IsFirstAllergenCheked();
                menuDayViewPage.ClickPlanView();

                menuDayViewPage.ClickPlanView();
                menuDayViewPage.ClickOnKeyword();
                var iskeywordChecked = menuDayViewPage.CheckKeywordchecked(keyword);
                menuDayViewPage.ClickPlanView();

                menuDayViewPage.ClickPlanView();
                menuDayViewPage.ClickOnKiosque();
                var iskiosqueChecked = menuDayViewPage.IsFirstkiosqueCheked();

                //Assert
                Assert.IsTrue(IsallergenChecked, "Les informations allergens n est pas enregistré.");
                Assert.IsTrue(iskeywordChecked, "Les informations keywords n est pas enregistré.");
                Assert.IsTrue(iskiosqueChecked, "Les informations kiosques n est pas enregistré.");
            }
            finally
            {
                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);

            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_GeneralInfoDeactivate()
        {
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string serviceMAD = TestContext.Properties["MenuService"].ToString();
            string infoDeactiveShouldBeDisplayed = "The menu cannot be deactivated.";
            string infoActiveShouldBeDisplayed = "Activated";
            var homePage = LogInAsAdmin();
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.Site, siteMAD);
            menusPage.Filter(MenusPage.FilterType.Service, serviceMAD);
            var menusDayViewPage = menusPage.SelectFirstMenu();
            MenusGeneralInformationPage menusGeneralInformation = menusDayViewPage.ClickOnGeneralInformation();
            menusGeneralInformation.WaitLoading();
            string infoDesactivate = menusGeneralInformation.InfoDeactivate().Trim();
            Assert.AreEqual(infoDesactivate, infoDeactiveShouldBeDisplayed, "Un menu ayant un service associé  doit apparaître bien en bas avec l'information :'The menu cannot be deactivated.'");
            menusDayViewPage = menusGeneralInformation.ClearService();
            menusGeneralInformation = menusDayViewPage.ClickOnGeneralInformation();
            string infoDesactivateAfterClearService = menusGeneralInformation.InfoDeactivate().Trim();
            Assert.AreEqual(infoDesactivateAfterClearService, infoActiveShouldBeDisplayed, "Un menu n'ayant pas de service associé doit apparaître en bas avec l'information :'Activated'");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_MehtodMassiveChangeConfirm()
        {
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string recipeName = "RecipeForTestMenu";
            string recipeNameBis = "Recipe2ForTestMenu";
            string method = "Std";
            
            var homePage = LogInAsAdmin();
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menuDayViewPage.AddRecipe(recipeName);
            menuDayViewPage.AddRecipe(recipeNameBis);
            menuDayViewPage.ChangeRecipeMethod(method);
            var isValid = menuDayViewPage.CheckMethodChanges(method);
            Assert.IsTrue(isValid, "La valeur de la methode n'a pas été chnagé");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_Kiosk()
        {
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string recipeName = "RecipeForTestMenu";
            string concept = "CH - Brioch";
            var date = DateTime.Now.ToString("dd/MM/yyyy",CultureInfo.InvariantCulture) +" - "+ DateTime.Now.ToString("dddd", CultureInfo.InvariantCulture); 

            var homePage = LogInAsAdmin();

            MenusPage menusPage = homePage.GoToMenus_Menus();
            try
            {
                MenusCreateModalPage menusCreateModalPage = menusPage.MenuCreatePage();
                MenusDayViewPage menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                MenusGeneralInformationPage generalInformationPage = menuDayViewPage.ClickOnGeneralInformation();
                generalInformationPage.AddConcept();
                generalInformationPage.ChooseMenuDay(date);
                WebDriver.Navigate().Refresh();
                menuDayViewPage.AddRecipe(recipeName);
                menuDayViewPage.ClickPlanView();
                menuDayViewPage.ClickOnKioskImage();
                string selectedKiosk = menuDayViewPage.SelectKiosk();
                var displayedKiosk = menuDayViewPage.VerifyKioskName(selectedKiosk);
                Assert.IsTrue(displayedKiosk, "It's not the same name");
            }
            finally
            {
                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_MethodEdit()
        {
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string recipeName = "Recipe2ForTestMenu";
            string recipeNameBis = "RecipeForTestMenu";
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeMethod = "Std";

            //Arrange
            
            var homePage = LogInAsAdmin();

            AddServiceForMenu(homePage, serviceName, guestType);
            var menusPage = homePage.GoToMenus_Menus();
            try
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menuDayViewPage.AddRecipe(recipeName);
                menuDayViewPage.AddRecipe(recipeNameBis);

                var pax_before_change_method = menuDayViewPage.GetPAX();
                var prixtotal_before_change_method = menuDayViewPage.GetPRIXXTOTAL();
                menuDayViewPage.ClickOnFirstRecipe();
                var initMethod = menuDayViewPage.GetRecipeMethod();
                menuDayViewPage.SetRecipeCoef("0");
                menuDayViewPage.SetRecipeMethod(recipeMethod);
                menuDayViewPage.SetRecipeCoef("0");

                WebDriver.Navigate().Refresh();

                var pax_after_change_method = menuDayViewPage.GetPAX();
                var prixtotal_after_change_method = menuDayViewPage.GetPRIXXTOTAL();

                Assert.AreNotEqual(pax_before_change_method, pax_after_change_method, "Le prix dans la colonne PAX a été mis à jour.");
                Assert.AreNotEqual(prixtotal_before_change_method, prixtotal_after_change_method, "Le prix total la colonne Sales Price a été mis à jour.");
            }
            finally
            {
                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);
            }


        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_DisplayViewNameChange()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            //string recipeName = "RecipeForTestMenu";
            string recipeName = "Rec";
            string recipeNameToUpdate = $"{recipeName}-updated" + new Random().Next().ToString();
            //Arrange
            var homePage = LogInAsAdmin();
            var menusPage = homePage.GoToMenus_Menus();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant);
            menusPage = menuDayViewPage.BackToList();

            //Assert
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);

            var menusDayViewPage = menusPage.SelectFirstMenu();
            menusDayViewPage.AddRecipe("Rec");
            //Get Recipe Name on Plan View before update on Display view 
            var originalRecipeBeforeUpdatedOnDisplayView = menusDayViewPage.GetRecipeName();
            menusDayViewPage.ClickPlanView();
            var recipeNameOnDisplayView = menusDayViewPage.GetRecipeNameOnDisplayView();
            menusDayViewPage.ClickOnFirstRecipeOnDispalyView();
            menusDayViewPage.SetFirstRecipeOnDispalyView(recipeNameToUpdate);
            var recipeNameUpdatedOnDisplayView = menusDayViewPage.GetRecipeNameOnDisplayView();
            menusDayViewPage.ClickPlanView();
            //Get Recipe Name on Plan View After updated on Display view 
            var originalRecipeAfterUpdatedOnDisplayView = menusDayViewPage.GetRecipeName();
            Assert.AreNotEqual(recipeNameOnDisplayView, recipeNameUpdatedOnDisplayView, "Le changement de nom de recette dans le display view a echoué");
            Assert.AreEqual(originalRecipeBeforeUpdatedOnDisplayView, originalRecipeAfterUpdatedOnDisplayView, " Le nom de recette dans le Plan View aprés le changement dans le Display View n'est pas le meme");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_index_CreateDayOfWeek()
        {

            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            DateTime StartDate = new DateTime(2024, 7, 07);
            DateTime EndDate = new DateTime(2024, 7, 23);

            //Arrange
        
            HomePage homePage = LogInAsAdmin();
           

            //AddServiceForMenu(homePage, serviceName, guestType);
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            //var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            var daySelected = menusCreateModalPage.SelectDays();
            menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            //homePage.Navigate();

            //AddServiceForMenu(homePage, serviceName, guestType);
            menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            var selectedMenuDayView = menusPage.SelectFirstMenu();
            var selectedMenuWeekView = selectedMenuDayView.SwitchToWeekView();
            Assert.IsTrue(selectedMenuDayView.CheckDaysDispoAreDaysSelected(daySelected), "les jours disponibles ne sont pas que ceux qui ont été sélectionné lors de la création");
            Assert.IsTrue(selectedMenuWeekView.CheckUnselectedDaysGris(daySelected), "les jours non sélectionné ne sont pas grisés");
            Assert.IsTrue(selectedMenuWeekView.CheckUnselectedHaveRightIcon(daySelected), "les jours non sélectionné ont icone Rond barré");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_DisplayViewKeyword()
        {
            Random rnd = new Random();
            string recipeName = recipeNameToday + "-" + rnd.Next().ToString();
            string menuName = menuNameToday + "-" + rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string keyword = "keyword_for_test";
            string ingredient = "itemForKeywordTest" + "-" + rnd.Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient, group, workshop, taxType, prodUnit);

            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            ItemKeywordPage itemKeywordPage = itemGeneralInformationPage.ClickOnKeywordItem();
            itemKeywordPage.AddKeyword(keyword);
            var isAdded = itemKeywordPage.IsKeywordAdded(keyword);
            Assert.IsTrue(isAdded, "keyword n'a pas ajouté");
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");
            recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList(); 
            recipeVariantPage.AddIngredient(ingredient);    

            AddServiceForMenu(homePage, serviceName, guestType);
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            try
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menuDayViewPage.AddRecipe(recipeName);
                menuDayViewPage.ClickPlanView();
                menuDayViewPage.ClickOnFirstRecipeOnDispalyView();
                menuDayViewPage.ClickOnKeyword();
                var iskeywordChecked = menuDayViewPage.CheckKeywordchecked(keyword);
                Assert.IsTrue(iskeywordChecked, "Les informations keywords n est pas enregistré.");

            }
            finally
            {
                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, ingredient);
                ItemGeneralInformationPage itemGeneralInformation = itemPage.ClickFirstItem();
                itemGeneralInformation.ClickOnDeleteItem();
                itemGeneralInformation.ConfirmDelete();
                itemPage = itemGeneralInformation.BackToList();
                MassiveDeleteModal massiveDeleteModal = itemPage.MenuMassiveDelete();
                massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.SearchByName, ingredient);
                massiveDeleteModal.ClickSearch();
                massiveDeleteModal.SelectFirst();
                massiveDeleteModal.DeleteFirstService();
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_CoeffMassiveChangeConfirm()
        {

            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeName = "RecipeForTestMenu";
            string recipeNameBis = "Recipe2ForTestMenu";
            string messageconfirmation = "You're about to change all the quantity for the menu day with the selected value. Do you confirm?";
            //Arrange
            var homePage = LogInAsAdmin();
            AddServiceForMenu(homePage, serviceName, guestType);
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menuDayViewPage.AddRecipe(recipeName);
            menuDayViewPage.AddRecipe(recipeNameBis);
            menuDayViewPage.SetSetCoeffMassiveCoef("14");
            menuDayViewPage.ClickValidatee();
            Assert.IsTrue(messageconfirmation.Equals(menuDayViewPage.GetMessageConfirmationMassive()), "le pop-up n'apparait pas pour demander confirmation ou annulation du Massive update du coeff..");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_CoeffMassiveChangeConfirms()
        {

            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeName = "RecipeForTestMenu";
            string recipeNameBis = "Recipe2ForTestMenu";
            //Arrange
            var homePage = LogInAsAdmin();
            AddServiceForMenu(homePage, serviceName, guestType);
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menuDayViewPage.AddRecipe(recipeName);
            menuDayViewPage.AddRecipe(recipeNameBis);
            menuDayViewPage.SetSetCoeffMassiveCoef("14");
            menuDayViewPage.ClickValidatee();
            menuDayViewPage.ClickConfirmationMassive();
            var CoefMassive = menuDayViewPage.GetAllCoefMassive();
            Assert.AreEqual(2, CoefMassive.Count(x => x == "14"), "Le coefficient de toutes les recettes ne sont  pas met à jour");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_DisplayViewAllergen()
        {
            Random rnd = new Random();
            string recipeName = recipeNameToday + "-" + rnd.Next().ToString();
            string menuName = menuNameToday + "-" + rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string allergen = TestContext.Properties["RecipeAllergen"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");

                recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                var itemIntolerancePage = recipeGeneralInfosPage.ClickOnIntolerancePage();
                itemIntolerancePage.AddFirstAllergen(allergen);

            }
            AddServiceForMenu(homePage, serviceName, guestType);
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
            menuDayViewPage.AddRecipe(recipeName);
            menuDayViewPage.ClickPlanView();
            menuDayViewPage.ClickOnAllergenEtoile();
            //Assert
            Assert.AreEqual("Lupin", menuDayViewPage.CheckDefaultAllergenchecked(), "Les allergens de la recette par défaut ne sont pas cochés");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_details_MethodMassiveChangeConfirm()
        {
            //Prepare
            string methodTest = "Std";
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string recipeName = "RecipeForTestMenu";
            string recipeName2 = "Recipe2ForTestMenu";
            Random rnd = new Random();
            string menuName = "MenuForMassiveChange" + rnd.Next();

            var homePage = LogInAsAdmin();
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            try
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayView = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);
                menuDayView.AddRecipe(recipeName);
                menuDayView.AddRecipe(recipeName2);
                menuDayView.ChangeRecipeMethod(methodTest);
                menuDayView.ClickValidatee();
                menuDayView.ClickConfirmationMassive();
                bool allMethod = menuDayView.GetAllMethod().All(method => methodTest == method);
                //Assert
                Assert.IsTrue(allMethod, "Les méthodes de toutes les recettes ne sont pas mises à jour");
            }
            finally
            {
                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_organizebyMenuName()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();

            menuMassiveDeletePage.MassiveDeleteSiteSearch(site);
            menuMassiveDeletePage.MenuNameTriParClick();

            //Assert
            Assert.IsTrue(menuMassiveDeletePage.IsSortedByName(), "Les Menu Name ne sont pas triés par ordre alphabétique.");
        }

        [TestMethod]

        [Timeout(_timeout)]
        public void RE_MENU_MassiveDelete_OrganizeByFrom()
        {
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.ClickOnSearch();
            menuMassiveDeletePage.ClickOnFrom();
            Assert.IsTrue(menuMassiveDeletePage.CheckOrderFrom(), "Order est decroissant");
            menuMassiveDeletePage.ClickOnFrom();
            Assert.IsFalse(menuMassiveDeletePage.CheckOrderFrom(), "Order est croissant");



        }


        [TestMethod]

        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_select()
        {

            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();

            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant);
            menusPage = menuDayViewPage.BackToList();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.Filter(MenuMassiveDeletePage.FilterType.SearchMenu, menuName);
            menuMassiveDeletePage.ClickOnSearch();
            menuMassiveDeletePage.ClickOnFirstLine();
            menuMassiveDeletePage.ClickOnDelete();
            menuMassiveDeletePage.ClickOnConfirmDelete();
            menuMassiveDeletePage.ClickOnConfirmAfterDelete();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            //Assert
            Assert.AreEqual(0, menusPage.CheckTotalNumber(), "Menu supprimé avec succès");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_FromToSearch()

        {
            DateTime DateFrom = DateTime.Today.AddMonths(-1);
            DateTime DateTo = DateTime.Today.AddDays(1);
            string dateFromString = DateFrom.ToString("dd/MM/yyyy");
            string dateToString = DateTo.ToString("dd/MM/yyyy");
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            MenusPage menusPage = homePage.GoToMenus_Menus();
            var recipeMassiveDelete = menusPage.OpenMassiveDeletePopup();
            menusPage.Fill(MenusPage.FillType.DateFrom, DateFrom);
            menusPage.Fill(MenusPage.FillType.DateTo, DateTo);
            menusPage.ClickSearch();
            var dateFromlist = menusPage.GetListDateFrom();
            var dateToList = menusPage.GetListDateTo();
            bool allDatesfromValid = menusPage.IsDateFromGreaterOrEqualToAll(dateFromlist, DateFrom);
            Assert.IsTrue(allDatesfromValid, $"le filtre date frim n'est pas fonctionnel .");
            bool allDatestoValid = menusPage.IsDateToGreaterOrEqualToAll(dateFromlist, DateTo);
            Assert.IsTrue(allDatestoValid, "le filtre dateto nest pas fonctionnel.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_link()
        {
            HomePage homePage =LogInAsAdmin();
           
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            MenusMassiveDeletePopup menuMassiveDeletePage = menusPage.MenuOpenMassiveDeletePopup();
            menuMassiveDeletePage.SelectAllSites();
            menuMassiveDeletePage.SelectAllVariant();
            menuMassiveDeletePage.ClickOnSearch();
            var MenuGeneralInformation = menuMassiveDeletePage.ClickMenuLinkFromRow(1);
            var windowIsOpened = menuMassiveDeletePage.VerifyIfNewWindowIsOpened();
            Assert.IsTrue(windowIsOpened, "Une nouvelle fenêtre n'a pas été ouverte comme prévu.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_status_activeMenu()
        {
            var ActivemenuStatus = "Active menu";
            var InactivemenuStatus = "Inactive menu";

            var homePage = LogInAsAdmin();

            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.Filter(MenuMassiveDeletePage.FilterType.Status, ActivemenuStatus);
            menuMassiveDeletePage.ClickOnSearch();
            menuMassiveDeletePage.ClickonSelectAll();
            var activemenusNumber = menuMassiveDeletePage.CheckSelectedMenus();
            menuMassiveDeletePage.ClickonUnSelectAll();
            menuMassiveDeletePage.Filter(MenuMassiveDeletePage.FilterType.Status, InactivemenuStatus);
            menuMassiveDeletePage.ClickOnSearch();
            menuMassiveDeletePage.ClickonSelectAll();
            var inactivemenusNumber = menuMassiveDeletePage.CheckSelectedMenus();
            Assert.AreNotEqual(activemenusNumber, inactivemenusNumber, "Les éléments actifs du menu ne sont pas filtrés.");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_selectall()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();

            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant);
            menusPage = menuDayViewPage.BackToList();
            menusPage.ResetFilter();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.Filter(MenuMassiveDeletePage.FilterType.SearchMenu, menuNameToday);
            menuMassiveDeletePage.ClickOnSearch();
            int nombredesligness = menuMassiveDeletePage.CountRows();
            menuMassiveDeletePage.ClickonSelectAll();
            menuMassiveDeletePage.ClickOnDelete();
            menuMassiveDeletePage.ClickOnConfirmDelete();
            menuMassiveDeletePage.ConfirmDeleteAll();
            MenuMassiveDeletePage menuMassiveDeletePages = menusPage.GoToMassiveDelete();
            menuMassiveDeletePages.Filter(MenuMassiveDeletePage.FilterType.SearchMenu, menuNameToday);
            menuMassiveDeletePages.ClickOnSearch();
            int nombredeslignes = menuMassiveDeletePages.CountRows() - 1;

            //Assert
            Assert.AreEqual(0, nombredeslignes, " Select all de massive delete ne fonctionne pas.");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_Variantsearch()
        {
            //Prepare
            string variantSearched = TestContext.Properties["MenuVariantACE1"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            var menuMassiveDeletePage = menusPage.MenuOpenMassiveDeletePopup();
            menuMassiveDeletePage.SelectVariantByName(variantSearched, false);
            menuMassiveDeletePage.ClickOnSearch();
            List<string> nameListe = menuMassiveDeletePage.GetAllNamesResultPaged();
            bool variantNameExistInAll = nameListe.All(variantName => variantSearched == variantName);
            //Assert
            Assert.IsTrue(variantNameExistInAll, "Tous les noms dans les résultats de recherche ne correspondent pas à la variante recherchée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_pagination()
        {

            var homePage = LogInAsAdmin();

            var menusPage = homePage.GoToMenus_Menus();

            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.ClickOnSearch();
            bool verifpagination = menuMassiveDeletePage.Pagination();

            int nombrelignesbefore = menuMassiveDeletePage.CountRows();
            menuMassiveDeletePage.ChangeNumberPagination();
            int nombrelignesafter = menuMassiveDeletePage.CountRows();

            if (nombrelignesbefore != nombrelignesafter && verifpagination == true)
            {
                bool verif = true;
                Assert.IsTrue(verif, "Test Validé");

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_organizebyTo()
        {
            var homePage = LogInAsAdmin();

            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.ClickOnSearch();
            menuMassiveDeletePage.ClickOnTo();
            Assert.IsTrue(menuMassiveDeletePage.CheckOrderToASC(), "Order est decroissant ");
            menuMassiveDeletePage.ClickOnTo();
            Assert.IsTrue(menuMassiveDeletePage.CheckOrderToDsc(), "Order est croissant ");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_status_inactiveMenus()
        {
            var MenuStatus = "Inactive menu";

            var homePage = LogInAsAdmin();
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.Filter(MenuMassiveDeletePage.FilterType.Status, MenuStatus);
            menuMassiveDeletePage.ClickOnSearch();

            var checkInactiveRow = menuMassiveDeletePage.CheckAllRowsAreInactive();


            Assert.IsTrue(checkInactiveRow, "Les résultats ne sont pas mis à jour en fonction des filtres MENU INACTIVE");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_organizebysite()
        {
            var MenuStatus = "Active menu";

            var homePage = LogInAsAdmin();
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.Filter(MenuMassiveDeletePage.FilterType.Status, MenuStatus);
            menuMassiveDeletePage.ClickOnSearch();
            menuMassiveDeletePage.SortBySite();
            var isSorted = menuMassiveDeletePage.IsSorted();
            Assert.IsTrue(isSorted, "Le tri par site ne fonctionne pas!");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_inactivesitessearch()
        {

            HomePage homePage= LogInAsAdmin();
            
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.CheckAllSites();
            menuMassiveDeletePage.ClickOnSearch();
            menuMassiveDeletePage.ClickonSelectAll();
            var activeSitemenusNumber = menuMassiveDeletePage.CheckSelectedMenus();
            menuMassiveDeletePage.ClickonUnSelectAll();
            menuMassiveDeletePage.ClickOnInactiveSiteCheck();
            menuMassiveDeletePage.SelectAllInactiveSites();
            menuMassiveDeletePage.ClickOnSearch();
            menuMassiveDeletePage.ClickonSelectAll();
            var inactiveSitemenusNumber = menuMassiveDeletePage.CheckSelectedMenus();
            Assert.AreNotEqual(activeSitemenusNumber, inactiveSitemenusNumber, "Les éléments inactifs du site ne sont pas filtrés.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_Menu_Display_View()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string recipeName = "RecipeForTestMenu";

            //Arrange
            var homePage = LogInAsAdmin();
            var menusPage = homePage.GoToMenus_Menus();

            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant);
            menusPage = menuDayViewPage.BackToList();

            //Assert
            menusPage.ResetFilter();
            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);

            var menusDayViewPage = menusPage.SelectFirstMenu();
            menusDayViewPage.AddRecipe(recipeName);
            //Get Recipe Name on Plan View before update on Display view 
            var RecipeName = menusDayViewPage.GetRecipeName();
            menusDayViewPage.ClickPlanView();

            menusDayViewPage.ClickOnFirstRecipeOnDispalyView();
            menusDayViewPage.ClickPlanView();

            bool IsRecipeNameOFF = menusDayViewPage.ClickOnFirstRecipeOFFDispalyView();

            Assert.IsTrue(IsRecipeNameOFF, "En désactivant Display View la recette doit apparaître sans avoir à réactualiser la page.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_WeekView_Print()
        {
            //Prepare
            string recipeName = "RecipeForTestMenu";
            string site = TestContext.Properties["Site"].ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string recipeSalePrice = "5";
            string recipePoints = "1";
            string TLSEnergyKCal = "600";
            //--------------------------------------------------------------------------------------
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            bool versionPrint = true;
            string DocFileNamePdfBegin = "Menu Per Week";
            string DocFileNameZipBegin = "All_files_";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            DateTime date = DateUtils.Now;
            List<string> listDateOfWeek = new List<string> { "IsActiveOnTuesday", "IsActiveOnThursday", "IsActiveOnSaturday" };
            //Arrange

            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");

                recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipeVariantPage.SetSalesPrice(recipeSalePrice);
                recipeVariantPage.SetPointsPrice(recipePoints);
                recipeVariantPage.ClickOnDieteticTab();
                if (!recipeVariantPage.GetEnergyKCal().Equals(TLSEnergyKCal))
                {
                    recipeVariantPage.SetEnergyKCal(TLSEnergyKCal);
                }
                recipesPage = recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInformationPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInformationPage.ClickOnFirstVariant();
                recipeVariantPage.SetSalesPrice(recipeSalePrice);
                recipeVariantPage.SetPointsPrice(recipePoints);
                recipeVariantPage.ClickOnDieteticTab();
                if (!recipeVariantPage.GetEnergyKCal().Equals(TLSEnergyKCal))
                {
                    recipeVariantPage.SetEnergyKCal(TLSEnergyKCal);
                }
                recipesPage = recipeVariantPage.BackToList();
            }
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            menusCreateModalPage.SelectManyDays(listDateOfWeek);
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant);
            var selectedMenuWeekView = menuDayViewPage.SwitchToWeekView();
            selectedMenuWeekView.ClearDownloads();
            try
            {
                Assert.IsTrue(selectedMenuWeekView.IsWeekViewVisible(), "La vue 'Week vue' n'est pas accessible sur l'écran.");
                int nb_days_selected = selectedMenuWeekView.AddRecipeForManyDatetime(recipeName, date);
                PrintReportPage printReportPage = selectedMenuWeekView.Print(versionPrint);
                bool isGenerated = printReportPage.IsReportGenerated();
                printReportPage.Close();
                Assert.IsTrue(isGenerated, "Le menu n'a pas été imprimé.");
                printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                // cliquer sur All
                string trouve = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                Assert.IsTrue(fi.Exists, trouve + " non généré");
                PdfDocument document = PdfDocument.Open(fi.FullName);
                Page page1 = document.GetPage(1);
                IEnumerable<Word> words = page1.GetWords();
                var nb_recipeName = words.Where(w => w.Text.Contains(recipeName)).Count();
                //Assert
                Assert.AreEqual(nb_days_selected, nb_recipeName, " les colonnes vides (Jours inectif) non présent dans le Pdf");
            }
            finally
            {
                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);
            }
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_DisplayView_AffihageBouton()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            try
            {
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant);

                menuDayViewPage.ClickPlanView();

                bool IsDisplayView = menuDayViewPage.MessageIsDisplayView();
                Assert.IsTrue(IsDisplayView, "Le bouton Switch ne fonctionne pas !.");

                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
            }
            finally
            {
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
                menusPage.MassiveDeleteMenus(menuName, site, variant);
                menusPage.ResetFilter();
            }
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_OrganizeByMainVariant()
        {
            //Prepare
            string site = TestContext.Properties["SiteACE"].ToString();

            //Arrange 
            HomePage homePage= LogInAsAdmin();
            
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.ResetFilter();

            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.MassiveDeleteSiteSearch(site);

            menuMassiveDeletePage.MenuVariantTriParClick();

            var sortBy = menuMassiveDeletePage.IsSortedByVariant();

            Assert.IsTrue(sortBy, "Les variants ne sont pas triés par ordre alphabétique pour le site: " + site);
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_ExportVariants()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string recipeName = "RecipeForTestMenu";
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            Random rnd = new Random();

            string menuName = menuNameToday + "-" + rnd.Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            MenusDayViewPage menuDayViewPage;

            //Arrange
            HomePage homePage = LogInAsAdmin();
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");

                recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
            }
            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant, serviceName);

            menuDayViewPage.AddRecipe(recipeName);

            var menuGeneralInformationPage = menuDayViewPage.ClickOnGeneralInformation();
            string Variant = menuDayViewPage.GetMenuVariant();
            menuDayViewPage.BackToList();

            menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
            //menusPage.ClickFirstMenu();

            menusPage.ClearDownloads();

            MenusWeekViewPage menusWeekViewPage = new MenusWeekViewPage(WebDriver, TestContext);
            menusWeekViewPage.ExportMenus();

            // Récupérer le fichier exporté
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = menusPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", filePath);
            var listResult = OpenXmlExcel.GetValuesInList("Variant", "Sheet 1", filePath);
            //Assert
            Assert.IsTrue(listResult.Contains(variant), MessageErreur.EXCEL_DONNEES_KO);
            Assert.AreNotEqual(listResult, Variant, "le fichier d'export contient les variants définis sur le menu");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_PopUpIcones()
        {
            //Prepare
            string site = "ACE";
            string variant = "ADULTOS 300G - ALMUERZO";
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.SelectFirstVariant();
            recipeGeneralInfosPage.Validate();
            recipesPage = recipeGeneralInfosPage.BackToList();
            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant);
            menuDayViewPage.AddRecipe(recipeName);
            menuDayViewPage.ClickOnDisplayView();
            menuDayViewPage.ClickOnFirstMenuRecipe();
            // Assert
            bool kiosquesClickResult = menuDayViewPage.ClickOnKiosquesIcon();
            bool allergensClickResult = menuDayViewPage.ClickOnAllergensIcon();
            bool keywordsClickResult = menuDayViewPage.ClickOnKeywordsIcon();

            // Validate results
            Assert.IsTrue(kiosquesClickResult, "L'icône Kiosques a bougé après le clic.");
            Assert.IsTrue(allergensClickResult, "L'icône Allergens a bougé après le clic..");
            Assert.IsTrue(keywordsClickResult, "L'icône Keywords a bougé après le clic.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Duplcation_Concept()
        {

            //Arrange      
            HomePage homePage = LogInAsAdmin();
            //Act
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.MenuCreatePage();
            var concept = menusPage.IsConceptUnique();

            //verifier si le champ concept est dupliqué
            Assert.IsTrue(concept, "The 'Concept' field is duplicated.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_PlanView_Switch_Button()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddMonths(1);
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string nbPortions = "5";
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string serviceName = TestContext.Properties["MenuService"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            try
            {
                // Create a recipe
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions);
                recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);
                AddServiceForMenu(homePage, serviceName, guestType);
                // Create a Menu
                var menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, startDate, endDate, site, variant);
                // add Recipe
                menuDayViewPage.AddRecipe(recipeName);
                menusPage = menuDayViewPage.BackToList();
                menusPage.ResetFilter();
                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
                var firstMenu = menusPage.SelectFirstMenu();  
                firstMenu.ClickOnDisplayView();
                bool isDisplayView = firstMenu.MessageIsDisplayView();
                Assert.IsTrue(isDisplayView, "le button de display view ne marche pas");
                List<string> listOfDisplayPage = firstMenu.GetTheFirstLineOfDisplayPage();

                firstMenu.ClickOnDisplayView();
                bool isPlanView = firstMenu.MessageIsPlanView();
                Assert.IsTrue(isPlanView, "le button de display view ne marche pas");
                List<string> listOfPlanPage = firstMenu.GetTheFirstLineOfPlanPage();
                Assert.AreNotEqual(listOfPlanPage, listOfDisplayPage, "le button de display view ne marche pas");
            }
            finally
            {
                //delete menu
                var menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);
                // delete recipe
                RecipesPage recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_OngletClinic()
        {
            // prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddMonths(1);
            bool isClinic = true ;

            // arrange
            HomePage homePage = LogInAsAdmin();

            // act
            var menusPage = homePage.GoToMenus_Menus();
            try
            {
                //create menu
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, startDate, endDate, site, variant);
                var generalInformation = menuDayViewPage.ClickOnGeneralInformation();

                generalInformation.SetClinic(isClinic);
                bool verif = generalInformation.IsClinicOpen();

                //Assert
                Assert.IsTrue(verif, "onglet de clinic n'apparait pas");
                isClinic = false;
                generalInformation.SetClinic(false);
                verif= generalInformation.IsClinicOpen();
                Assert.IsFalse(verif, "onglet de clinic apparait");
            }
            finally
            {
                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Export()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string recipeName = "RecipeForTestMenu";
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            Random rnd = new Random();

            string menuName = "Menu-" + DateUtils.Now.ToString("dd/MM/yyyy") + rnd.Next().ToString();
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            var dateFrom = DateUtils.Now;
            var dateTo = DateUtils.Now.AddMonths(1);

            MenusDayViewPage menuDayViewPage;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            // Go to menu Recipes and filter with name recipe
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, "5");

                recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
            }
            //Act
            AddServiceForMenu(homePage, serviceName, guestType);

            // Go to Menu Menus
            var menusPage = homePage.GoToMenus_Menus();
            try
            {
                // Create New Menu and add Recipe
                var menusCreateModalPage = menusPage.MenuCreatePage();
                menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, dateFrom, dateTo, site, variant, serviceName);

                menuDayViewPage.AddRecipe(recipeName);
                menusPage = menuDayViewPage.BackToList();

                menusPage.Filter(MenusPage.FilterType.SearchMenu, menuName);
                menusPage.Filter(MenusPage.FilterType.DateFrom, dateFrom);
                menusPage.Filter(MenusPage.FilterType.DateTo, dateTo);

                // Clear dowloads and export Menu filtred
                menusPage.ClearDownloads();
                menusPage.ExportMenus();

                // Récupérer le fichier exporté
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();
                var correctDownloadedFile = menusPage.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile);
                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);
                int resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", filePath);
                var ResultName = OpenXmlExcel.GetValuesInList("Menu", "Sheet 1", filePath);
                //Assert
                Assert.AreEqual(resultNumber, 1, "Tous les menus filtrés n' apparaitre pas dans l'export");
                Assert.IsTrue(ResultName.Contains(menuName), MessageErreur.EXCEL_DONNEES_KO);
            }
            finally
            {
                // Delete Menu Created
                menusPage.ClosePrintButton();
                menusPage.WaitPageLoading();
                menusPage.MassiveDeleteMenus(menuName, site, variant);

            }


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_MENU_Massivedelete_inactivesitescheck()
        {
            string site = "Inactive -";
            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            var menusPage = homePage.GoToMenus_Menus();
            MenuMassiveDeletePage menuMassiveDeletePage = menusPage.GoToMassiveDelete();
            menuMassiveDeletePage.ClickOnInactiveSiteCheck();
            var isAllInactive = menuMassiveDeletePage.SiteComboBoxChecker(site);
            // Assert
            Assert.IsTrue(isAllInactive, "Les Sites ne sont pas Inactive.");
        }
    }
}